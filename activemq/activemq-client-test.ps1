﻿#
# A powershell script to read from any ActiveMq provider
# 

[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$Username,
    [parameter(Mandatory=$true)][string]$Password
    )

Import-Module -Name PSActiveMQClient

# Local settings. Customize as needed
#$ActiveMQServer = "failover:(tcp://msgbroker1.tier2.sfu.ca:61616,tcp://msgbroker2.tier2.sfu.ca:61616)?randomize=false"
$ActiveMQServer = "activemq:tcp://localhost:61616"
$queueName = "ICAT.amaint.toExchange"
$retryQueueName = "ICAT.amaint.toExchange.retry"
$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"
# Maximum number of times to retry processing a message. After this message will be logged and discarded
$MaxRetries = 30 


## Local private functions ##

# In case of error, shut down ActiveMQ session and exit
function graceful-exit($s)
{
    try {
        Remove-ActiveMQSession $s
    }
    catch {
    }
    exit 0
}

# Process an ActiveMQ message from Amaint
# First see if user needs an Exchange mailbox. Lightweight & disabled accts don't
# Next, check if the user exists in AD. If not, skip this message - we have to wait for AD handler to create user
# If user exists, enable Exchange mailbox if necessary and then verify account settings
function process-amaint-message($xmlmsg)
{
    $username = $xmlmsg.synclogin.username

    # Skip lightweight and non-active accts
    if ($xmlmsg.syncLogin.login.isLightweight -eq "true" -or $xmlmsg.syncLogin.login.status -ne "active")
    {
        Add-Content $Logfile "$(date) : Skipping update for $username. Lightweight or inactive"
        return 1
    }

    # Verify the user in AD
    try {
        $aduser = Get-ADUser $username
    }
    catch {
        # Either they don't exist or there's an AD error. Either way we can't continue
        Add-Content $Logfile "$(date) : $username not found in AD. Failing: $_"
        return 0
    }
    
    $groups = $xmlmsg.synclogin.login.adGroups.childNodes.InnerText
    $aliases = $xmlmsg.synclogin.login.aliases.childNodes.InnerText
    $roles = $xmlmsg.synclogin.person.roles.childNodes.InnerText

    # For testing
    write-host "User     : $username"
    Write-host "AD Groups: $($groups -join ',')"
    Write-Host "Aliases  : $($aliases -join ',')"
    Write-Host "Roles    : $($roles -join ',')"
    write-host ""

}


# Queue a message in the retry queue to retry it later
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

    if ($mtmp.retryMessage.count -gt $MaxRetries)
    {
        Add-Content $Logfile "$(date) : FAIL. Max retries exceeded for " + $mtmp
        return 0
    }

    Send-ActiveMQMessage -Queue "queue://$retryQueueName" -Session $AMQSession -Message $mtmp

    # Ack the original message
    $m.Acknowledge()

    return 1

}

## end local functions



## main code block

$AMQSession = New-ActiveMQSession -Uri $ActiveMQServer -User $Username -Password $Password -ClientAcknowledge

$Target = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$queueName")
$RetryTarget = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$retryQueueName")

# Create a consumer with the target
$Consumer =  $AMQSession.CreateConsumer($Target)
$RetryConsumer = $AMQSession.CreateConsumer($RetryTarget)

# Wait for a message. For now, we'll wait a really short time and 
# if no message arrives, sleep before trying again. That way we can add more logic
# inside our loop later if we want to (e.g. checking multiple queues for messages)

$loopcounter=1;

while(1)
{
    try {
        $Message = $Consumer.Receive([System.TimeSpan]::FromTicks(10000))
        if (!$Message)
        {
            $loopcounter++
            if ($loopcounter > 10)
            {
                # Only try the retry queue every 10 seconds
                $Message = $RetryConsumer.Receive([System.TimeSpan]::FromTicks(10000))
                $loopcounter = 1
            }
            if (!$Message)
            {
                Start-Sleep -Seconds 1
                continue
            }
            [xml]$msgtmp = $Message.Text
            $msg = $msgtmp.retryMessage
            Add-Content $Logfile "$(date) : Retrying msg $msgtmp"
        }
        else
        {
            [xml]$msg = $Message.Text
        }

        # We currently only care about SyncLogin messages
        if ($msg.syncLogin)
        {
            Add-Content $Logfile "$(date) : Processing Amaint msg $msg"
            if (process-amaint-message($msg))
            {
                Add-Content $Logfile "$(date) : Success"
                $Message.Acknowledge()
            }
            else
            {
                Add-Content $Logfile "$(date) : Failure. Will Retry"
                retry-message($Message)
            }
        }
        # Add 'if' blocks here for other message types
        else
        {
            Add-Content $Logfile "$(date) : Ignoring msg: $msg"
            $Message.Acknowledge()
        }
    }
    catch {
        $_
        # Realistically, we want to log errors but try to recover
        # For now we'll just exit and let Windows Scheduler restart us
        graceful-exit($AMQSession)
    }
}


