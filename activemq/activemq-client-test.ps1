#
# A powershell script to read from any ActiveMq provider
# 

[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$Username,
    [parameter(Mandatory=$true)][string]$Password
    )

Import-Module -Name PSActiveMQClient

function graceful-exit($s)
{
    try {
        Remove-ActiveMQSession $s
    }
    catch {
    }
    exit 0
}

function process-message($xmlmsg)
{
    $username = $xmlmsg.synclogin.username
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




#$ActiveMQServer = "failover:(tcp://msgbroker1.tier2.sfu.ca:61616,tcp://msgbroker2.tier2.sfu.ca:61616)?randomize=false"
$ActiveMQServer = "activemq:tcp://localhost:61616"
$queueName = "ICAT.amaint.toExchange"
$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"


$AMQSession = New-ActiveMQSession -Uri $ActiveMQServer -User $Username -Password $Password -ClientAcknowledge

$Target = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$queueName")

# Create a consumer with the target
$Consumer =  $AMQSession.CreateConsumer($Target)

# Wait for a message. For now, we'll wait a really short time and 
# if no message arrives, sleep before trying again. That way we can add more logic
# inside our loop later if we want to (e.g. checking multiple queues for messages)

while(1)
{
    try {
        $Message = $Consumer.Receive([System.TimeSpan]::FromTicks(10000))
        if (!$Message)
        {
            Start-Sleep -Seconds 1
            continue
        }
        [xml]$msg = $Message.Text
        # We only care about SyncLogin messages
        if (!$msg.syncLogin)
        {
            # Acknowledge but skip
            $msg.Acknowledge()
            Add-Content $Logfile "$(date) : Ignoring msg: $msg"
            continue
        }
        process-message($msg)
        $Message.Acknowledge()
    }
    catch {
        $_
        # Realistically, we want to log errors but try to recover
        # For now we'll just exit and let Windows Scheduler restart us
        graceful-exit($AMQSession)
    }
}


