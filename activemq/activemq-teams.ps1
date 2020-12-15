# ActiveMQ client for MS Teams management
# 
# This client listens for AMQ messages sent by the Request-A-Team Webapp
# It only handles creating teams and will also update owners and descriptions
# if it receives a request for a team that already exists
#
# This script is designed to be run periodically, rather than as a daemon.
# Invoke from the Task Scheduler and run at regular intervals, e.g. every 10 minutes
#
# Although it *could* be run as a service, Powershell scripts don't play nicely as 
# a Service, so a third party app is needed to manage the service. Not worth it for this

$me = $env:username
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\m365teams.log"

# Import modules we need
Import-Module -Name MicrosoftTeams -NoClobber
Import-Module -Name PSActiveMQClient

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:RestToken = $settings.RestToken
    $global:GroupsOU = $settings.AzureGroupsOU
    $global:TeamsAdmin = $settings.TeamsAdmin
    $global:TeamsAdminPW = $settings.TeamsAdminPW
    $global:CompositeAttribute = $settings.AzureGroupsCompositeAttribute

    $global:Username = $settings.amqUsername
    $global:Password = $settings.amqPassword
    $global:queueName = $settings.TeamsQueueName
    $global:retryQueueName = $settings.TeamsRetryQueueName
    $global:RestToken = $settings.RestToken
    $global:MaxRetries = $settings.MaxRetries
    $global:MaxRetryTimer = $settings.MaxRetryTimer
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:MaxNoActivity = $settings.MaxNoActivity
    $global:SmtpServer = $settings.SmtpServer
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}


function process-message($xmlmsg)
{
    if (-not $xmlmsg.teamsRequest )
    {
        # Not a teams request message. No others are supported yet
        return 1
    }

    $teamname = "SFU Teams - $($xmlmsg.teamsRequest.name)"
    $mailnickname = $teamname  -replace "[^A-Za-z0-9_-]",""
    $owners = $xmlmsg.teamsRequest.owners.Split(",")
    $descr = $xmlmsg.teamsRequest.description

    $hasMaillists = $false
    if ($xmlmsg.teamsRequest.maillists)
    {
        $hasMaillists = $true
        $maillists = $xmlmsg.teamsRequest.maillists.ChildNodes.InnerText -join ","
    }

    try {
        # Check for existing Team
        $team = Get-Team -DisplayName $teamname
        if ($team -eq $null)
        {
            # Create team
            $team = New-Team -DisplayName "$teamname" -Description "$descr" -Owner "$($owners[0])@sfu.ca" -ShowInTeamsSearchAndSuggestions $false `
                    -MailNickname $mailnickname
            Write-Log "Created new Team `"$teamname`" with GroupID $($team.GroupID)"
        }
    }
    catch {
        # If we get an error here, we can't really continue
        Write-Log "Error retrieving or creating Team $teamname: $($_.Exception)"
        return 0
    }

    try {
        # Add all owners, just in case
        foreach ($owner in $owners)
        {
            Add-TeamUser -GroupID $team.GroupID -user "$($owner)@sfu.ca" -role "Owner"
        }
    }
    catch {
        # If we get an error here, we can't really continue
        Write-Log "Error adding Owners to $teamname: $($_.Exception)"
        return 0
    }

    # Create the AD composite group if necessary
    if ($hasMaillists)
    {
        $create = $false
        try {
            $adgroup = GET-ADGroup -Identity "M365 $($xmlmsg.teamsRequest.name)" -Properties $CompositeAttribute
        } catch
        {
            $create = $true
        }
        try {
            if ($create -or $adgroup -eq $null)
            {
                New-ADGroup -Name "M365 $($xmlmsg.teamsRequest.name)" -GroupCategory Security -GroupScope Global `
                    -DisplayName "M365 $($xmlmsg.teamsRequest.name)" -Path $GroupsOU -Description "MAILLISTS=$maillists;TEAM=$($team.GroupdID)"
                
                Write-Log "Created AD Group `"M365 $($xmlmsg.teamsRequest.name)`""
            } 
            else
            {
                # Make sure the description is set correctly
                $adgroup | Set-ADGroup -Description "MAILLISTS=$maillists;TEAM=$($team.GroupID)"
            }
        }
        catch {
            # If we get an error here, we can't really continue
            Write-Log "Error creating AD Group for $teamname: $($_.Exception)"
            return 0
        }
    }
    return 1
}


# Queue a message in the retry queue to retry it later.
# We reformat the XML - wrap it in a "retryMessage" tag and
# add a retry count tag.
function retry-message($m)
{
    [xml]$mtmp = $m.Text
    $firstRetry = $false
    # Add a retry counter if one isn't already there
    if (! $mtmp.retryMessage.count)
    {
        # This is a bit clunky - couldn't find a good way to insert a
        # counter into the XML message so we'll create a new "retry" message type with a counter element
        [xml]$retrymsg = "<retryMessage><count>1</count>" + $mtmp.InnerXml + "</retryMessage>"
        $mtmp = $retrymsg
    }
    # Otherwise add one to the retry count
    else
    {
        $count = [int]$mtmp.retryMessage.count
        if ($count -eq 1)
        {
            $firstRetry = $true
        }
        $count++
        $mtmp.retryMessage.count = "$count"
    }

    if ([int]$mtmp.retryMessage.count -gt $MaxRetries)
    {
        Write-Log "FAIL. Max retries exceeded for $($mtmp.InnerXml)"
        return 0
    }

    Send-ActiveMQMessage -Queue $retryQueueName -Session $AMQSession -Message $mtmp

    if ($firstRetry)
    {
        return 2
    }
    return 1

}

## end local functions



## main code block

load-settings($SettingsFile)

# Set up our Exchange shell
$e_uri = $ExchangeServer + "/PowerShell/"
try {
    # Set up AzureAD session
    # get credentials and login as AAD admin
    $Password = $global:TeamsAdminPW | ConvertTo-SecureString -AsPlainText -Force
    $UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $global:TeamsAdmin,$Password

    $junk = Connect-MicrosoftTeams -Credential $UserCredential
}
catch {
        write-log "Error connecting to MS Teams: "
        write-log $_.Exception
        exit
}

Write-Log "Starting up"

$AMQSession = New-ActiveMQSession -Uri $ActiveMQServer -User $Username -Password $Password -ClientAcknowledge

$Target = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$queueName")
$RetryTarget = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$retryQueueName")

# Create a consumer with the target
$Consumer =  $AMQSession.CreateConsumer($Target)
$RetryConsumer = $AMQSession.CreateConsumer($RetryTarget)

# Wait for a message. For now, we'll wait a really short time and 
# if no message arrives, sleep before trying again. That way we can add more logic
# inside our loop later if we want to (e.g. checking multiple queues for messages)

$loopcounter=1
$noactivity=0
$retryTimer=10
$retryFailures=0
$msg=""

# Not running in daemon mode, don't loop forever. Run once and exit
#while(1)
{
    try {
        $isRetry = $false
        $Message = $Consumer.Receive([System.TimeSpan]::FromTicks(10000))
        if (!$Message)
        {
            # Because this script will only run periodically, there is no harm in
            # checking the retry queue every time we run, so skip the checks below.
            # Leaving the code in though in case we want to convert this to a continually running daemon

            # No message from the main queue. See if we should check the retry queue
            $loopcounter++
            # Only try the retry queue every x seconds, where x is 10x number of failures in a row
            # This acts as a sequential backoff timer. The retry queue is first tried after 10 seconds
            # then if a message in the retry queue fails again, the next retry time will be 20 seconds,
            # and so on. That way, if there's a failure in underlying infrastructure, the message should
            # still eventually get processed. 
            ##if ($loopcounter -gt $retryTimer)
            {
                $Message = $RetryConsumer.Receive([System.TimeSpan]::FromTicks(10000))
                # Only if no message was found, reset loop counter so that if there are 
                # multiple messages to be tried, they'll all be tried at once
                if (!$Message) 
                { 
                    $loopcounter = 1
                    # Also reset the number of retry Failures, since there's no msgs left
                    if ($retryFailures)
                    {
                        Write-Log "No messages in retry queue. Clearing retryFailures" 
                        $retryFailures = 0
                        $retryTimer=10
                    }
                }
                # Also reset the loop counter if last retrymsg failed, so that we back off properly.
                elseif ($retryFailures) { $loopcounter = 1 }
            }
            if (!$Message)
            {
                $noactivity++
                if ($noactivity -gt $MaxNoActivity)
                {
                    $noactivity=0
                    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "No activity from ActiveMQ for $MaxNoActivity seconds" `
                    -SmtpServer $SmtpServer -Body "Seems a bit fishy."
                }
                Start-Sleep -Seconds 1
                continue
            }

            # Got a message from the Retry queue. Extract the inner message
            $isRetry=$true
            # undef the msg variable before defining it, because retry msgs and regular msgs are slightly different object types
            Remove-Variable msg
            [xml]$msgtmp = $Message.Text
            $msg = $msgtmp.retryMessage
            Write-Log "Retrying msg `r`n$($msgtmp.InnerXml)"
        }
        else
        {
            # undef the msg variable before defining it, because retry msgs and regular msgs are slightly different object types
            Remove-Variable msg
            [xml]$msg = $Message.Text
        }

        $noactivity=0

        if (-Not $isRetry) { Write-Log "Processing msg `r`n $($msg.InnerXml)" }
        $rc = process-message($msg)
        Write-Log "RC = $rc"
        if ($rc -gt 0)
        {
            Write-Log "Success"
            $rc = $Message.Acknowledge()
            if ($isRetry) 
            { 
                $retryFailures = 0 
                $retryTimer=10
            }
        }
        else
        {
            if ($isRetry) 
            { 
                $retryFailures++ 
                $retryTimer = (1+$retryFailures) * 10
                if ($retryTimer -gt $MaxRetryTimer) { $retryTimer = $MaxRetryTimer }
                Write-Log "Retry backoff is now $retryTimer seconds"
            }
            Write-Log "Failure. Will Retry"
            $rc = retry-message($Message)

            if ($rc -eq 2)
            {
                # Special case: The first time processing a msg in the retry queue, if it fails, treat it as a success
                # for the purposes of calculating the backoff, so that if there are messages queued behind it, we get
                # through them all quickly.
                $retryFailures = 0
            }
            # Even if retry-message exceeds max retries, we still have to Acknowledge msg to clear it from the queue
            $Message.Acknowledge()
            if ($rc -eq 0)
            {
                Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "Failure from Exchange ActiveMQ handler" `
                    -SmtpServer $SmtpServer -Body "Failed to process message $MaxRetries time.`r`nMessage: $($Message.Text). `r`nLast Error: $LastError"
            }
        }

    }
    catch {
        $_
        # Realistically, we want to log errors but try to recover
        # For now we'll just exit and let Windows Scheduler restart us
        write-host "Caught error. Closing sessions"
        Remove-ActiveMQSession $AMQSession
        exit 0
    }
}

Remove-ActiveMQSession $AMQSession
exit 0
