#
# A powershell script to read from any ActiveMq provider
# 

[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$Username,
    [parameter(Mandatory=$true)][string]$Password
    )

Import-Module -Name PSActiveMQClient
Import-Module -Name PSAOBRestClient

# Local settings. Customize as needed
$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
#$ActiveMQServer = "failover:(tcp://msgbroker1.tier2.sfu.ca:61616,tcp://msgbroker2.tier2.sfu.ca:61616)?randomize=false"
$ActiveMQServer = "activemq:tcp://localhost:61616"
$queueName = "ICAT.amaint.toExchange"
$retryQueueName = "ICAT.amaint.toExchange.retry"
$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"
$TokenFile = "C:\Users\$me\REST_Authtoken.txt"

# Maximum number of times to retry processing a message. After this message will be logged and discarded
$MaxRetries = 30 
# Listname for users who are on Exchange. Ignore ActiveMQ msgs for anyone not on this list
$ExchangeUsersList = "exchange-users"


$RestToken = Get-Content $TokenFile -totalcount 1

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

## Local private functions ##

# In case of error, shut down ActiveMQ session and exit
function graceful-exit($s)
{
    try {
        Remove-ActiveMQSession $s
        Remove-PSSession $ESession
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
        # TODO: Add check here to disable accounts that go lightweight or disabled
        # Note that if we use Disable-Mailbox to disconnect an AD user from their mailbox,
        # we have to use connect-mailbox rather than enable-mailbox to reconnect them
        # It may be better to create a retention policy that deletes all mail after x days and
        # change the mailbox's retention policy to that, plus prevent logins by:
        #  - disabling all protocols for account
        #  - change email aliases to, e.g "_disabled"
        # for reference: Set-CASMailbox USER -ActiveSyncEnabled $false -ImapEnabled $false -EwsEnabled $false -MAPIEnabled $false -OWAEnabled $false -PopEnabled $false -OWAforDevicesEnabled $false
        # What if we just set AccountDisabled to $true with set-mailbox?
        # maybe disable-mailbox after account is inactive for 1(?) year?
        Add-Content $Logfile "$(date) : Skipping update for $username. Lightweight or inactive"
        return 1
    }

    # Skip users not on Exchange yet. Remove this check when all users are on.
    try {
        $rc = Get-AOBRestMaillistMembers -Maillist $ExchangeUsersList -Member $username -AuthToken $RestToken
    }
    catch {
        Add-Content $Logfile "$(date) : Error communicating with REST Server for $username. Aborting processing of msg. $_"
        return 0
    }

    if (-Not $rc)
    {
        Add-Content $Logfile "$(date) : Skipping update for $username. Not a member of $ExchangeUsersList"
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

    $create = $false
    $update = $false
    # See if the user already has an Exchange Mailbox
    try {
        $mb = Get-Mailbox $username
    }
    catch {
        # It's possible that other errors could trigger a failure here but we'll deal with that below
        $create = $true
        $update = $true
    }
    
    # No mailbox exists. Enable the mailbox in Exchange
    if ($create)
    {
        # TODO: We need to determine whether the user previously had an Exchange mailbox and
        # if so, use Connect-Mailbox to reconnect them, as Enable-Mailbox will always create a new mailbox.
        try {
            Enable-Mailbox -Identity $username -name $username
            $mb = Get-Mailbox $username
        }
        catch {
            # Now we have a problem. Throw an error and abort for this user
             Add-Content $Logfile "$(date) : Unable to enable Exchange Mailbox for ${username}: $_"
             return 0
        }
    }

    # Default to hidden in GAL
    $hideInGal=$true

    $roles = @($xmlmsg.synclogin.person.roles.InnerText)
    if ($roles -contains "staff" -or $roles -contains "faculty" -or [int]$xmlmsg.synclogin.person.sfuVisibility -gt 4)
    {
        $hideInGal=$false
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
            $update = $true
        }

        if ($mb.HiddenFromAddressListsEnabled -ne $hideInGal)
        {
            $update = $true
        }
    }

    if ($update)
    {
        # TODO: If there are any other attributes we should set on new or changed mailboxes, do it here
        $addresses = @($xmlmsg.synclogin.login.aliases.ChildNodes.InnerText) -Join ","
        try {
            Set-Mailbox -Identity $username -HiddenFromAddressListsEnabled $hideInGal -EmailAddresses $addresses
            Add-Content $Logfile "$(date) : Updated mailbox for ${username}. HideInGal: $hideInGal. Aliases: $addresses"
        }
        catch {
            Add-Content $Logfile "$(date) : Unable to update Exchange Mailbox for ${username}: $_"
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

    if ($mtmp.retryMessage.count -gt $MaxRetries)
    {
        Add-Content $Logfile "$(date) : FAIL. Max retries exceeded for $($mtmp.InnerXml)"
        return 0
    }

    Send-ActiveMQMessage -Queue $retryQueueName -Session $AMQSession -Message $mtmp

    # Ack the original message
    $rc = $m.Acknowledge()

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
            # No message from the main queue. See if we should check the retry queue
            $loopcounter++
            if ($loopcounter -gt 10)
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

            # Got a message from the Retry queue. Extract the inner message
            [xml]$msgtmp = $Message.Text
            $msg = $msgtmp.retryMessage
            Add-Content $Logfile "$(date) : Retrying msg `r`n$($msgtmp.InnerXml)"
        }
        else
        {
            [xml]$msg = $Message.Text
        }

        # We currently only care about SyncLogin messages
        if ($msg.syncLogin)
        {
            Add-Content $Logfile "$(date) : Processing Amaint msg `r`n $($msg.InnerXml)"
            if (process-amaint-message($msg))
            {
                Add-Content $Logfile "$(date) : Success"
                $rc = $Message.Acknowledge()
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
            $rc = $Message.Acknowledge()
        }
    }
    catch {
        $_
        # Realistically, we want to log errors but try to recover
        # For now we'll just exit and let Windows Scheduler restart us
        graceful-exit($AMQSession)
    }
}


