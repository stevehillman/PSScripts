﻿#
# A powershell script to read from any ActiveMq provider
# 

# Ensure that Exchange cmdlets throw a catchable error when they fail
$ErrorActionPreference = "Stop"


Import-Module -Name PSActiveMQClient
Import-Module -Name PSAOBRestClient

$me = $env:username
$LogFile = "C:\Users\$me\activemq_client.log"
$SettingsFile = "C:\Users\$me\settings.json"
$ConnectUsersFile = "C:\Users\$me\ConnectUsers.json"


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
    $global:ExchangeUsersListPrimary = $settings.ExchangeUsersListPrimary
    $global:ExchangeUsersListSecondary = $settings.ExchangeUsersListSecondary
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:MaxNoActivity = $settings.MaxNoActivity
    $global:SmtpServer = $settings.SmtpServer
    $global:AddNewUsersDate = $settings.AddNewUsersDate
    $global:AddAllUsersDate = $settings.AddAllUsersDate
    $global:PassiveMode = ($settings.PassiveMode -eq "true")
    $global:SubscribeURL = $settings.SubscribeURL
    $global:AddNewUsers = $false
    $global:AddAllUsers = $false

    $global:ExternalDomains = @()
    $settings.ExternalDomains | ForEach {
        $global:ExternalDomains += $_
    }
}

$global:ConnectUsersDate = "00000000"


# Load in the list of current Connect Users once a day. Relies on an external process
# to update the contents of the ConnectUsers file.
function Load-ConnectUsers()
{
    $global:now = Get-Date -Format FileDate
    Write-Log "now: $now"
    
    if ($global:ConnectUsersDate -lt $now -and (Test-Path $ConnectUsersFile))
    {
        Write-Log "ConnectUsersDate: $ConnectUsersDate, now: $now"
        $global:ConnectUsers = ConvertFrom-Json ((Get-Content $ConnectUsersFile) -join "")
        $global:ConnectUsersDate = $now
        Write-Log "Imported $($ConnectUsers.PSObject.Properties.Name.Count) users from $ConnectUsersFile"
    }
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

# Add user to a maillist.
# Returns $true on success, $false if they're already a member.
# Throws an exception if any other error occurred.
function Add-UserToMaillist($u,$l)
{
    $url = $SubscribeURL + $l + "&address=" + $u
    try {
        $result = Invoke-RestMethod -Method "GET" -Uri $url -ErrorAction 'Stop'
    }
    catch {
        # REST call failed. Bad news
        Write-Log "Failed to add $u to $l : $_"
        Throw "Failed to add member to list"
    }
    if ($result -match "^ok")
    {
        return $true
    } 
    elseif ($result -match "already a member")
    {
        return $false
    }
    else
    {
        # Something else went wrong
        Throw $result
    }
}

function process-message($xmlmsg)
{
    if ($msg.synclogin)
    {
        $global:now = Get-Date -Format FileDate
        if (-Not $AddNewUsers)
        {
            if ($now -ge $AddNewUsersDate)
            {
                $global:AddNewUsers = $true
                Write-Log("$AddNewUsersDate has passed. Invoking NewUser processing")
            }
        }
        if (-Not $AddAllUsers)
        {
            if ($now -ge $AddAllUsersDate)
            {
                $global:AddAllUsers = $true
                Write-Log("$AddAllUsersDate has passed. Invoking AllUser processing")

            }
        }
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

function send-compromisedresult($result, $status, $responseQueue)
{
    $elem = $result.CreateElement("statusMsg")
    $elem.InnerText = $status
    $junk = $result.compromisedlogin.AppendChild($elem)
    Send-ActiveMQMessage -Queue $responseQueue -Message $result -Session $AMQSession
}

function process-compromised-message($xmlmsg)
{
    $username = $xmlmsg.compromisedlogin.username

    # Build a Result object
    $result = [xml]"<compromisedLogin><messageType>response</messageType></compromisedLogin>"
    # Create a username element
    $elem = $result.CreateElement("username")
    $elem.InnerText = $username
    # Add the element to the XML object
    $junk = $result.compromisedlogin.AppendChild($elem)
    # Repeat for serial #
    $elem = $result.CreateElement("serial")
    $elem.InnerText = $xmlmsg.compromisedlogin.serial
    $junk = $result.compromisedlogin.AppendChild($elem)
    # And for app Name
    $elem = $result.CreateElement("application")
    $elem.InnerText = "Exchange"
    $junk = $result.compromisedlogin.AppendChild($elem)

    Write-Log "Processing Compromised Account $username"

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

    $scopedusername = $username + "@sfu.ca"

    # Verify the user has a mailbox
    try {
        $mailbox = Get-Mailbox $scopedusername -ErrorAction Stop
    }
    catch {
        # Nope
        Write-Log "No mailbox found for $scopedusername. Skipping"
        if ($xmlmsg.compromisedLogin.respond -and $xmlmsg.compromisedLogin.respond -ne "")
        {
            send-compromisedresult $result "No mailbox found for $username. Skipping" $xmlmsg.compromisedLogin.respond
        }
        return 1
    }

    # Unless compromised-message says otherwise, clear Exchange settings
    if ($xmlmsg.compromisedlogin.settings.resetEmailSettings -match "false")
    {
        return 1
    }

    # Start of account cleanup

    # For reference, here are the settings that were reset in Zimbra:
    # Reset all Zimbra settings that Spammers might monkey with. Refer to Confluence
    # for the authoritative list. So far it consists of
    #
    #  zmprov settings:
    #  zimbraPrefOutOfOfficeReply (default: "")
    #  zimbraPrefSaveToSent (default: TRUE)
    #  zimbraPrefAutoAddAddressEnabled (default: TRUE)
    #  zimbraMailSieveScript (parsed to disable any filter that forwards or discards mail)
    #
    # Note: the rest of the settings are per identity and multiple identities may exist
    # We just clear primary identity. Secondary identities will be ignored:
	# zimbraPrefFromDisplay
	# zimbraPrefFromAddress
	# zimbraPrefReplyToDisplay
	# zimbraPrefReplyToAddress
    # zimbraPrefReplyToEnabled
    # zimbraPrefMailSignature

    # Disable OWA and MAPI access. 
    # Is there any point in doing this? It doesn't kill existing sessions, just prevents new ones from being started
    # NOTE: There needs to be a way to re-enable this after a user has changed their password.
    #Set-CASMailbox $scopedusername -OWAEnabled $false -MAPIEnabled $false

    # Disable all rules
    try {
        $rules = Get-InboxRule -Mailbox $scopedusername | where {$_.Enabled  -and ($_.DeleteMessage -eq $true -or $_.ForwardTo -match "[a-z]+" -or $_.RedirectTo -match "[a-z]+")}
        $rules | Disable-InboxRule
        
        # When there's only one rule, $rules isn't an array, so $rules.count will be undefined
        $rulecount = 1;
        if ($rules.Count -eq 0 -or $rules.Count -gt 1)
        {
            $rulecount = $rules.Count
        }
        $response = "$rulecount rule(s) disabled. "
    }
    catch {
        $response = "An error occurred disabling rules: $_ . "
    }

    # Clear signatures
    try {
        Set-MailboxMessageConfiguration $scopedusername -SignatureHTML "" -SignatureText ""
    }
    catch {
        $response = $response + "Error clearing signatures: $_ . "
    }

    # Ensure mail isn't being forwarded. This should never be the case as this attribute is not
    # exposed to the user, but check anyway
    if ($mb.ForwardingSmtpAddress -and $mailbox.ForwardingSmtpAddress -ne "")
    {
        try {
            Set-Mailbox $scopedusername -DeliverToMailboxAndForward $false -ForwardingSmtpAddress ""
            $response = $response + "Forwarding cleared from $($mailbox.ForwardingSmtpAddress). "
        }
        catch {
            $response = $response + "Error clearing forwarding: $_ . "
        }
    }

    if ($xmlmsg.compromisedLogin.respond -and $xmlmsg.compromisedLogin.respond -ne "")
    {
        send-compromisedresult $result $response $xmlmsg.compromisedLogin.respond
    }
    return 1
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
        # Special case - ignore 'pending create' or 'defined' status (any others to ignore?)
        if ($xmlmsg.synclogin.login.status -eq "pending create" -or $xmlmsg.synclogin.login.status -eq "defined")
        {
            Write-Log "Skipping Pending Create or Defined status msg"
            return 1
        }
        $mbenabled = $false
        # TODO: Revisit how accounts get disabled. For now: 
        #  - prevent logins by disabling all protocols for account
        #  - change email aliases, appending "_disabled"
        #  - force HideInGal to True
        # 
        # maybe disable-mailbox after account is inactive for 1(?) year?
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
    $AddToMaillist = $false
    $AddToLightweightMigrations = $false

    $scopedusername = $username + "@sfu.ca"
    
    # See if the user already has an Exchange Mailbox
    try {
        $mb = Get-Mailbox $scopedusername -ErrorAction Stop
        $casmb = Get-CASMailbox $scopedusername -ErrorAction Stop
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

    $roles = @($xmlmsg.synclogin.person.roles.ChildNodes.InnerText)

    if ($create -and $AddNewUsers)
    {
        # See if ConnectUsers hash needs reloading
        Load-ConnectUsers

        # Check whether the user exists in Connect.
        if ($ConnectUsers.$username)
        {
            # From now (Aug 31/2018) on, treat all accounts that 
            # still exist in Connect as though they're lightweight transitioning
            # back to fullweight - i.e. create them new in Exchange, but flag that their
            # content from Connect needs migrating. We'll do that until our
            # migration software license expires
           
            $AddToLightweightMigrations = $true
            
        }
    }
    
    # No mailbox exists, Enable the mailbox in Exchange
    if ($create)
    {
        # TODO: We need to determine whether the user previously had an Exchange mailbox and
        # if so, use Connect-Mailbox to reconnect them, as Enable-Mailbox will always create a new mailbox.
        # For now, we never disable a mailbox, we just disable access to it.
        Write-Log "Creating mailbox for $username"
        if ($PassiveMode)
        {
            Write-Log "PassiveMode: Enable-Mailbox -Identity $username"
            # Simulate what a get-mailbox call would return
            $mb = New-Object -TypeName PSObject
            Add-Member -InputObject $mb EmailAddresses @("$($username)@sfu.ca")
            Add-Member -InputObject $mb HiddenFromAddressListsEnabled $true
            $casmb = New-Object -TypeName PSObject
            Add-Member -InputObject $casmb OWAEnabled $true
        }
        else 
        {
            try {
                $junk = Enable-Mailbox -Identity $scopedusername -ErrorAction Stop
                $junk = Set-CASMailbox $scopedusername -PopEnabled $false -OwaMailboxPolicy "Default" -ErrorAction Stop
                $mb = Get-Mailbox $scopedusername -ErrorAction Stop
            }
            catch {
                # Now we have a problem. Throw an error and abort for this user
                 $global:LastError = "Unable to enable Exchange Mailbox for ${username}: $_"
                 Write-Log $LastError
                 return 0
            }
#            try {
#                Set-MailboxMessageConfiguration $scopedusername -IsReplyAllTheDefaultResponse $false -ErrorAction Stop
#            }
#            catch {
#                Write-Log "Unable to set default OWA settings for $scopedusername, but safe to continue"
#            }

          
            if ($AddToLightweightMigrations)
            {
                try {
                    $junk = Add-UserToMaillist $username "lightweight-migrations" 
                    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "Lightweight acct $username converted to Exchange" `
                    -SmtpServer $SmtpServer -Body "Make sure their SFUConnect data gets migrated."
                }
                catch {
                    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "Error adding $username to list lightweight-migrations" `
                    -SmtpServer $SmtpServer -Body "$_"
                }
            }
        }
    }

    # Default to hidden in GAL
    $hideInGal=$true

    if ($roles -contains "staff" -or $roles -contains "faculty" -or ($roles -contains "other" -and [int]$xmlmsg.synclogin.person.sfuVisibility -gt 4))
    {
        if ($mbenabled)
        {
            $hideInGal=$false
        }
    }

    # Save for later comparisons
    $PreferredEmail = $xmlmsg.synclogin.person.email

    $AmaintAliases = @($xmlmsg.syncLogin.login.aliases.ChildNodes.InnerText)

    # Determine whether we should omit user's aliases from Exchange. For most Student
    # accounts, their alias is their name, which is personal information, so if we 
    # limit their account info to just their account name, we minimize PII exposed in Exchange.
    # The AD update scripts will ensure their names are masked if their sfuVisibility is < 5
    $maskAliases = $false
    if ($hideInGal -and [int]$xmlmsg.synclogin.person.sfuVisibility -lt 5)
    {
        $maskAliases = $true
        $PreferredEmail = $scopedusername
        $AmaintAliases = @()
    }



    if ($mbenabled -ne $casmb.OWAEnabled)
    {
        Write-Log "Account status changed. Updating"
        $update=$true
    }

    # Check if the account needs updating
    if (! $update)
    {
        #### Check aliases ####
        # Get the list of aliases from Exchange
        $al_tmp = @($mb.EmailAddresses)
        # Create empty array to hold unscoped aliases
        $aliases = @()

        $x = 0
        foreach ($alias in $al_tmp)
        {
            # Strip Exchange prefix and domain suffixes
            if ($alias -cmatch "SMTP")
            {
                $primaryaddress = $alias -replace "SMTP:"
            }
            $a = $alias  -replace ".*:" -replace "@.*"
            if ($a -ne $username -and $aliases -notcontains $a)
            {
                # Only add aliases once and that aren't the user's computing ID
                $aliases += $a
            }
        }   
        # compare-object returns non-zero results if the arrays aren't identical. That's all we care about
        try {
            if (Compare-Object -ReferenceObject $aliases -DifferenceObject $AmaintAliases)
            {
                Write-Log "Aliases have changed. Exchange had: $($aliases -join ','). Updating"
                $update = $true
            }
        }
        catch {
            # The above can fail if the user has no aliases. Set update to true *just in case*
            $update = $true
        }


        #### Check Primary SMTP Address ####
        if (!($primaryaddress -eq $PreferredEmail) -and $primaryaddress -Notmatch "\+sfu_connect")
        {
            # Primary SMTP address doesn't match, force update
            Write-Log "$primaryaddress doesn't match $PreferredEmail. Updating"
            $update = $true
        }

        #### Check GAL visibility ####
        if ($mb.HiddenFromAddressListsEnabled -ne $hideInGal)
        {
            Write-Log "HideInGal state changed. Updating"
            $update = $true
        }

        ## Check Display Name if appropriate
        if ($roles -contains "staff" -or $roles -contains "faculty" -or $roles -contains "other" -or [int]$xmlmsg.synclogin.person.sfuVisibility -gt 4)
        {
            $anondn = $false
            if ($mb.DisplayName -ne $xmlmsg.synclogin.login.gcos)
            {
                Write-Log "DisplayName doesn't match. Updating"
                $update = $true
            }
        }
        else 
        {
            $anondn = $true
            if ($mb.DisplayName -ne $username)
            {
                Write-Log "DisplayName is not anonymized and needs to be. Updating"
                $update = $true
            }
        }
    }

    if ($AddNewUsers -and $mb.PrimarySmtpAddress -Match "\+sfu_connect")
    {
        # Once all new users go into Exchange, process every account EXCEPT accounts
        # that were imported from Zimbra but haven't been migrated yet
        $update = $false
    }

    if ($update)
    {
        # TODO: If there are any other attributes we should set on new or changed mailboxes, do it here

        # For that rare case when a user has specified a non-SFU PreferredEmail address in SFUDS
        if ($PreferredEmail -Notmatch "@.*sfu.ca")
        {
            # See if the domain of the PreferredEmail is whitelisted
            if ($PreferredEmail -match "@(.*)")
            {
                $domain = $Matches[1]
                if ($ExternalDomains -notcontains $domain)
                {
                    Write-Log "User's PreferredEmail domain $domain is not whitelisted. Setting to default address"
                    $PreferredEmail = $username + "@sfu.ca"
                }
                else {
                    Write-Log "User's PreferredEmail domain $domain is whitelisted. Allowing it"
                }
            }
            else 
            {    
                $PreferredEmail = $username + "@sfu.ca"
            }
        }

        $addresses = @($username) + $AmaintAliases
        # $addresses will contain duplicates because PreferredEmail is always going to be one of the aliases or the username
        # We'll deal with that below

        ## Security check - if PreferredEmail is an @*sfu.ca address, make sure its one of the user's own addresses
        if ($PreferredEmail -match "@.*sfu.ca" -and $addresses -notcontains ($PreferredEmail -replace "@.*sfu.ca"))
        {
            Write-Log "WARNING: $scopedusername Preferred Email address $PreferredEmail is not one of the their aliases. Ignoring"
            $PreferredEmail = $scopedusername
        }

        # Preferred address comes first in EmailAddresses list
        $addresses = @($PreferredEmail) + $addresses

        $ScopedAddresses = @()
        if ($mbenabled)
        {
            $primaryemail = $PreferredEmail
            ForEach ($addr in $addresses) {
                if ($addr -Notmatch "@")
                {
                    $Scopedaddr = $addr + "@sfu.ca"
                }
                else 
                {
                    $Scopedaddr = $addr
                }
                if ($ScopedAddresses -contains $Scopedaddr)
                {
                    # eliminate duplicates
                    continue
                }
                if ($Scopedaddr -ne "@sfu.ca")
                {
                    $ScopedAddresses += $Scopedaddr
                }
            }
            if ($anondn)
            {
                $DisplayName = $username
            }
            else 
            {
                $DisplayName = $xmlmsg.synclogin.login.gcos    
            }
        }
        else 
        {
            $primaryemail = $username + "_disabled@sfu.ca"
            $scopedaddresses += $primaryemail
            $DisplayName = $username
        }

        try {
            if ($PassiveMode)
            {
                Write-Log "PassiveMode: Set-Mailbox -Identity $scopedusername -HideInGal $hideInGal -EmailAddresses $ScopedAddresses"
            }
            else 
            {
                $junk = Set-Mailbox -Identity $scopedusername -HiddenFromAddressListsEnabled $hideInGal `
                            -EmailAddressPolicyEnabled $false `
                            -EmailAddresses $ScopedAddresses `
                            -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update `
                            -DisplayName $DisplayName `
                            -ErrorAction Stop
                Write-Log "Updated mailbox for ${scopedusername}. HideInGal: $hideInGal. Aliases: $ScopedAddresses; DisplayName: $DisplayName"
                $junk = Set-MailboxMessageConfiguration $scopedusername -IsReplyAllTheDefaultResponse $false -ErrorAction Stop

            }

            if ($mbenabled -ne $casmb.OWAEnabled)
            {
                if ($PassiveMode)
                {
                    Write-Log "PassiveMode: Set-CASMailbox $username -Enabled $mbenabled"
                }
                else 
                {    
                    Write-Log "Setting Account-Enabled state to $mbenabled"
                    $junk = Set-Mailbox -Identity $scopedusername -PrimarySmtpAddress $primaryemail -ErrorAction Stop
                    $junk = Set-CASMailbox $scopedusername -ActiveSyncEnabled $mbenabled `
                                        -ImapEnabled $mbenabled `
                                        -EwsEnabled $mbenabled `
                                        -MAPIEnabled $mbenabled `
                                        -OWAEnabled $mbenabled `
                                        -OWAforDevicesEnabled $mbenabled `
                                        -ErrorAction Stop
                    if (-not $mbenabled)
                    {
                        $junk = Set-CASMailbox $scopedusername -PopEnabled $false -ErrorAction Stop
                    }
                }
            }
            
        }
        catch {
            $global:LastError =  "Unable to update Exchange Mailbox for ${username}: $_"
            Write-Log $LastError
            return 0
        }
    }
    Write-Log "Got here. Returning Success"

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


