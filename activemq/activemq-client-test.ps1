#
# A powershell script to read from any ActiveMq provider
# 

Import-Module -Name PSActiveMQClient
Import-Module -Name PSAOBRestClient

$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"
$SettingsFile = "C:\Users\$me\settings.json"


## Local private functions ##

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:ActiveMQServer = $settings.ActiveMQServer
    $global:Username = $settings.amqUsername
    $global:Password = $settings.amqPassword
    $global:queueName = $settings.QueueName
    $global:retryQueueName = $settings.RetryQueueName
    $global:RestToken = $settings.RestToken
    $global:MaxRetries = $settings.MaxRetries
    $global:MaxRetryTimer = $settings.MaxRetryTimer
    $global:ExchangeUsersList = $settings.ExchangeUsersList
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
    if ($msg.synclogin)
    {
        return process-amaint-message($xmlmsg)
    }
    # Add other message types here in the future
    else
    {
        Write-Log "Ignoring msg: Unsupported type"
        return 1
    }
}

# Process an ActiveMQ message from Amaint
# First see if user needs an Exchange mailbox. Lightweight & disabled accts don't
# Next, check if the user exists in AD. If not, skip this message - we have to wait for AD handler to create user
# If user exists, enable Exchange mailbox if necessary and then verify account settings
$global:LastError=""
function process-amaint-message($xmlmsg)
{
    $username = $xmlmsg.synclogin.username

    # Skip lightweight and non-active accts
    $mbenabled = $true
    if ($xmlmsg.syncLogin.login.isLightweight -eq "true" -or $xmlmsg.syncLogin.login.status -ne "active")
    {
        $mbenabled = $false
        # TODO: Revisit how accounts get disabled. For now: 
        #  - prevent logins by disabling all protocols for account
        #  - change email aliases, appending "_disabled"
        #  - force HideInGal to True
        # 
        # maybe disable-mailbox after account is inactive for 1(?) year?
    }

    # Skip users not on Exchange yet. Remove this check when all users are on.
    try {
        $rc = Get-AOBRestMaillistMembers -Maillist $ExchangeUsersList -Member $username -AuthToken $RestToken
    }
    catch {
        $global:LastError =  "Error communicating with REST Server for $username. Aborting processing of msg. $_"
        Write-Log $LastError
        return 0
    }

    if (-Not $rc)
    {
        Write-Log "Skipping update for $username. Not a member of $ExchangeUsersList"
        return 1
    }

    Write-Log "Processing update for $username"

    # Verify the user in AD
    try {
        $aduser = Get-ADUser $username
    }
    catch {
        # Either they don't exist or there's an AD error. Either way we can't continue
        $global:LastError = "$username not found in AD. Failing: $_"
        Write-Log $LastError
        return 0
    }

    $create = $false
    $update = $false
    # See if the user already has an Exchange Mailbox
    try {
        $mb = Get-Mailbox $username
        $casmb = Get-CASMailbox $username
    }
    catch {
        # It's possible that other errors could trigger a failure here but we'll deal with that below
        if (-Not $mbenabled)
        {
            Write-Log "$username disabled or lightweight and has no Exchange Mailbox. Skipping"
            return 1
        }    
        $create = $true
        $update = $true
    }
    
    # No mailbox exists, Enable the mailbox in Exchange
    if ($create)
    {
        # TODO: We need to determine whether the user previously had an Exchange mailbox and
        # if so, use Connect-Mailbox to reconnect them, as Enable-Mailbox will always create a new mailbox.
        Write-Log "Creating mailbox for $username"
        try {
            Enable-Mailbox -Identity $username -name $username
            $mb = Get-Mailbox $username
        }
        catch {
            # Now we have a problem. Throw an error and abort for this user
             $global:LastError = "Unable to enable Exchange Mailbox for ${username}: $_"
             Write-Log $LastError
             return 0
        }
    }

    # Default to hidden in GAL
    $hideInGal=$true

    $roles = @($xmlmsg.synclogin.person.roles.InnerText)
    if ($roles -contains "staff" -or $roles -contains "faculty" -or ($roles -contains "other" -and [int]$xmlmsg.synclogin.person.sfuVisibility -gt 4))
    {
        if ($mbenabled)
        {
            $hideInGal=$false
        }
    }

    if ($mbenabled -ne $casmb.OWAEnabled)
    {
        Write-Log "Account status changed. Updating"
        $update=$true
    }

    # Check if the account needs updating
    if (! $update)
    {
        # Check aliases
        # Get the list of aliases from Exchange
        # Strip Exchange prefix and domain suffix
        $al_tmp = @($mb.EmailAddresses)
        # Create empty array of appropriate size to hold scoped aliases
        $aliases = @($null) * $al_tmp.count

        $x = 0
        foreach ($alias in $al_tmp)
        {
            $aliases[$x] = $alias  -replace ".*:" -replace "@.*"
            $x++
        }   
        # compare-object returns non-zero results if the arrays aren't identical. That's all we care about
        if (Compare-Object -ReferenceObject $aliases -DifferenceObject @($xmlmsg.syncLogin.login.aliases.ChildNodes.InnerText))
        {
            Write-Log "Aliases have changed. Exchange had: $($aliases -join ','). Updating"
            $update = $true
        }

        if ($mb.HiddenFromAddressListsEnabled -ne $hideInGal)
        {
            Write-Log "HideInGal state changed. Updating"
            $update = $true
        }
    }

    if ($update)
    {
        # TODO: If there are any other attributes we should set on new or changed mailboxes, do it here
        $addresses = @($xmlmsg.synclogin.login.aliases.ChildNodes.InnerText)
        if ($mbenabled)
        {
            $addresses = $addresses | % { $_ + "@sfu.ca"}
        }
        else 
        {
            $addresses = $addresses | % { $_ + "_disabled@sfu.ca"}
        }

        try {
            Set-Mailbox -Identity $username -HiddenFromAddressListsEnabled $hideInGal -EmailAddresses $addresses
            Write-Log "Updated mailbox for ${username}. HideInGal: $hideInGal. Aliases: $addresses"
            if ($mbenabled -ne $casmb.OWAEnabled)
            {
                Write-Log "Setting Account-Enabled state to $mbenabled"
                Set-CASMailbox $username -ActiveSyncEnabled $mbenabled `
                                        -ImapEnabled $mbenabled `
                                        -EwsEnabled $mbenabled `
                                        -MAPIEnabled $mbenabled `
                                        -OWAEnabled $mbenabled `
                                        -PopEnabled $mbenabled `
                                        -OWAforDevicesEnabled $mbenabled
            }
        }
        catch {
            $global:LastError =  "Unable to update Exchange Mailbox for ${username}: $_"
            Write-Log $LastError
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
        $count++
        $mtmp.retryMessage.count = "$count"
    }

    if ([int]$mtmp.retryMessage.count -gt $MaxRetries)
    {
        Write-Log "FAIL. Max retries exceeded for $($mtmp.InnerXml)"
        return 0
    }

    Send-ActiveMQMessage -Queue $retryQueueName -Session $AMQSession -Message $mtmp

    return 1

}

## end local functions



## main code block

load-settings($SettingsFile)

# Set up our Exchange shell
$e_uri = $ExchangeServer + "/PowerShell/"
try {
        if ($me -eq "hillman")
        {
            # For testing..
            $Cred = Get-Credential
            $ESession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $e_uri  -Authentication Kerberos -Credential $Cred
        }
        else
        {
            # Prod
            $ESession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $e_uri  -Authentication Kerberos
        }
        import-pssession $ESession
}
catch {
        write-host "Error connecting to Exchange Server: "
        write-host $_.Exception
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
        if (process-message($msg))
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


