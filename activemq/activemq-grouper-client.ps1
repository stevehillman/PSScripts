#
# A powershell script to read from any ActiveMq provider
# 

# Ensure that Exchange cmdlets throw a catchable error when they fail
$ErrorActionPreference = "Stop"


Import-Module -Name PSActiveMQClient
Import-Module -Name PSAOBRestClient

$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"
$SettingsFile = "C:\Users\$me\settings.json"


## Local private functions ##

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ActiveMQServer = $settings.ActiveMQServer
    $global:Username = $settings.amqUsername
    $global:Password = $settings.amqPassword
    $global:queueName = $settings.GrouperQueueName
    $global:retryQueueName = $settings.GrouperRetryQueueName
    $global:RestToken = $settings.RestToken
    $global:MaxRetries = $settings.MaxRetries
    $global:MaxRetryTimer = $settings.MaxRetryTimer
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:MaxNoActivity = $settings.MaxNoActivity
    $global:SmtpServer = $settings.SmtpServer
    $global:PassiveMode = ($settings.PassiveMode -eq "true")
    
}

$global:ExcludedUsersDate = "00000000"


function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

function process-message($xmlmsg)
{
    if ($msg.synclogin)
    {
        $global:now = Get-Date -Format FileDate
        return process-amaint-message($xmlmsg)
    }
    elseif ($msg.compromisedlogin)
    {
        return process-compromised-message($xmlmsg)
    }
    # Add other message types here in the future
    else
    {
        Write-Log "Ignoring msg: Unsupported type"
        return 1
    }
}

# Compare two arrays for differences and return a Custom Object with two properties - an array of OnlyInOne and an array of OnlyInTwo
# To achieve performance, we use the LINQ MS Framework. This is complete overkill for a few hundred members, but for lists with many 
# thousands, it becomes significant. Code was shamelessly stolen from Stack Overflow
#
# https://stackoverflow.com/questions/6368386/comparing-two-arrays-get-the-values-which-are-not-common
#
# See here for more info on using LINQ with Powershell: https://www.red-gate.com/simple-talk/dotnet/net-framework/high-performance-powershell-linq/

# This function has not yet been tested

function compare-arrays($arrayobj1,$arrayobj2)
{
    # First, Cast as Strings, just in case they aren't
    [string[]]$array1 = $arrayobj1
    [string[]]$array2 = $arrayobj2
    $onlyInOne = [string[]][Linq.Enumerable]::Except($array1, $array2)
    $onlyInTwo = [string[]][Linq.Enumerable]::Except($array2, $array1)
    $value = "" | Select-Object -Property OnlyInOne, OnlyInTwo
    $value.OnlyInOne = $onlyInOne
    $value.OnlyInTwo = $onlyInTwo

    return $value
}

# Process a Grouper ActiveMQ message.
# At SFU, we've got custom ActiveMQ code that runs on the AMQ server and "squashes" Grouper event messages
# (which consist of individual add/drop member events) down into a single "AD Update" message. These messages
# are also rate-limited to 1 per minute, to ensure AD doesn't get overloaded. So if 100 users are added to a list,
# only a single AD Update message will be sent.
#
# However, the AMQ code then does 3 things:
# - calls the AmaintRest API to determine all AD groups that have this group as a member
# - ForEach group, fetch all members (also done via AmaintRest), and write the members to a file named for the group name
# - FTP each file to a Windows server (tecnically this step is done using a Camel route, rather than our custom code)
#
# The old Windows DOS scripts would then process each of those files, except that it wasn't, as Alan found it didn't work for him. He
# was relying on GDD instead
#
# As such, this code has nothing to process yet, as we don't want to continue the process of using FTP to move data
#
# In order to simplify the AMQ code, we should consider stripping it down to a more basic level of functionality - just
# send the group names, once per minute, in AMQ messages. If we do that, the below code then needs to:
# - decode the message (Grouper messages are JSON, not XML, so it's a bit different)
# - Fetch parent groups via AmaintRest
# - Use AOBRest to fetch the members of each (I think AmaintRest does not flatten the memberships, so isn't what we want here)
# - compare the membership with what is currently in AD
# - apply changes in chunks



function process-grouper-message($xmlmsg)
{
    # Parse JSON message
    $GroupName = 

    # Fetch the parent groups OR start supporting nested groups. If we support nested groups, parent groups would have this group as a member
    # so they wouldn't need updating
    $ParentGroups = call_to_amaint_rest_to_Get_parent_groups($GroupName)

    ForEach ($Pgroup in $ParentGroups) {
        # Do we need to format PGroup as a DN for the ADGroup commands? Needs testing

        # If it's an add or update, test to see whether group exists
        try {
            $ADGroup = Get-ADGroup $PGroup -ErrorAction Stop
        }
        catch {
            # Group doesn't exist (maybe?). Create it
            Write-Log "Group $PGroup doesn't exist. Attempting to create it"
            try {
                Add-ADGroup $PGroup - Syntax for specifying OU? -ErrorAction Stop
                Write-Log "Group $PGroup created"
            }
            catch {
                Oh Oh. Cant continue.. Log error and return failure.
            }
        }
        # Fetch memberships
        $ADGroupMembers = Get-ADGroupMembers $PGroup
        # These probably then need processing into a simple array of sAMAccountNames that can be compared against Amaint
        $MLGroupMembers = Get-AOBRestMaillistMembers -Maillist $PGroup -AuthToken $RestToken
        $AddsDrops = compare-arrays($ADGroupMembers,$MLGroupMembers)
        $toAdd = $AddsDrops.OnlyInTwo
        $toDrop = $AddsDrops.OnlyInOne

        if ($ToAdd)
        {
            $n = [System.Math]::Ceiling( ($ToAdd.Count / 1000) )
            $Chunks = Split-Array -Array $ToAdd -Chunks $n
            Write-Log "Adding $($toAdd.Count) members in $n chunks of max 1000"
            
            foreach ($Chunk in $Chunks)
            {
                try { 
                    Add-ADGroupMember -Identity $PGroup -Members $Chunk  -ErrorAction Stop 
                }
                catch { 
                    If this fails, cant continue. Log error and return failure.
                }
            }
        }
        if ($toDrop)
        {
            $n = [System.Math]::Ceiling( ($toDrop.Count / 1000) )
            $Chunks = Split-Array -Array $toDrop -Chunks $n
            Write-Log "Removing $($toDrop.Count) members in $n chunks of max 1000"

            
            foreach ($Chunk in $Chunks)
            {
                try { 
                    Add-ADGroupMember -Identity $PGroup -Members $Chunk  -ErrorAction Stop 
                }
                catch { 
                    If this fails, cant continue. Log error and return failure.
                }
            }
        }
    }
    # If we got here, we succeeded
    Write-Log "Group $GroupName successfully processed"
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

while(1)
{
    try {
        $isRetry = $false
        $Message = $Consumer.Receive([System.TimeSpan]::FromTicks(10000))
        if (!$Message)
        {
            # No message from the main queue. See if we should check the retry queue
            $loopcounter++
            # Only try the retry queue every x seconds, where x is 10x number of failures in a row
            # This acts as a sequential backoff timer. The retry queue is first tried after 10 seconds
            # then if a message in the retry queue fails again, the next retry time will be 20 seconds,
            # and so on. That way, if there's a failure in underlying infrastructure, the message should
            # still eventually get processed. 
            if ($loopcounter -gt $retryTimer)
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
        Remove-PSSession $ESession
        Remove-ActiveMQSession $AMQSession
        exit 0
    }
}


