#
# A powershell script to read from any ActiveMq provider
# 

# Ensure that Exchange cmdlets throw a catchable error when they fail
$ErrorActionPreference = "Stop"


Import-Module -Name PSActiveMQClient

$me = $env:username
$LogFile = "C:\Users\$me\activemq_amaint_client.log"
$SettingsFile = "C:\Users\$me\settings.json"
$cipherfile = "C:\user_updates\amaintcipherkey"

## The Chilkat DLL is needed for Blowfish decryption. It's a commercial license, to
## also needs to be "unlocked" at startup
## For details: https://www.chilkatsoft.com/refdoc/csCrypt2Ref.html
Add-Type -Path ".\lib\ChilkatDotNet48.dll"

## Unlock code taken from: https://www.example-code.com/powershell/global_unlock.asp
# The Chilkat API can be unlocked for a fully-functional 30-day trial by passing any
# string to the UnlockBundle method.  A program can unlock once at the start. Once unlocked,
# all subsequently instantiated objects are created in the unlocked state. 
# 
# After licensing Chilkat, replace the "Anything for 30-day trial" with the purchased unlock code.
# To verify the purchased unlock code was recognized, examine the contents of the LastErrorText
# property after unlocking.  For example:
$glob = New-Object Chilkat.Global
$success = $glob.UnlockBundle("Anything for 30-day trial")
if ($success -ne $true) {
    $($glob.LastErrorText)
    exit
}

$status = $glob.UnlockStatus
if ($status -eq 2) {
    $("Unlocked using purchased unlock code.")
}
else {
    $("Unlocked in trial mode.")
}

# Define the Blowfish object
$crypt = New-Object Chilkat.Crypt2


## Local private functions ##

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ActiveMQServer = $settings.ActiveMQServer
    $global:Username = $settings.amqUsername
    $global:Password = $settings.amqPassword
    $global:queueName = $settings.AmaintQueueName
    $global:retryQueueName = $settings.AmaintRetryQueueName
    $global:MaxRetries = $settings.MaxRetries
    $global:MaxRetryTimer = $settings.MaxRetryTimer
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:MaxNoActivity = $settings.MaxNoActivity
    $global:SmtpServer = $settings.SmtpServer
    $global:PassiveMode = ($settings.AmaintPassiveMode -eq "true")
    $global:cipherkey = [System.IO.File]::ReadAllBytes($cipherfile)    
    $global:testusers = @('ebronte','kipling','ebrontst','kiptest')
    $global:grouproles = @('staff','faculty','grad','undergrad')
    $global:UsersOU = $settings.UsersOU
    $global:GroupsOU = $settings.GroupsOU
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

# Process an Amaint ActiveMQ message.
# The logic as as follows:
# - If the account is defined or pending create, skip - it has never existed in AD
# - If the account is disabled or destroyed, disable it in AD if it exists
# - if the account is active, create or update it (we always process every update, since we can't tell whether the password has changed)

function process-amaint-message($xmlmsg)
{
    $username = $xmlmsg.synclogin.username
    $scopedusername = $username + "@sfu.ca"

    $passive = ($global:PassiveMode -and -not ($global:testusers -contains $username))

    # Skip lightweight and non-active accts
    $userenabled = $true
    if ($xmlmsg.syncLogin.login.status -ne "active")
    {
        # Special case - ignore 'pending create' or 'defined' status (any others to ignore?)
        if ($xmlmsg.synclogin.login.status -eq "pending create" -or $xmlmsg.synclogin.login.status -eq "defined")
        {
            Write-Log "Skipping Pending Create or Defined status msg"
            return 1
        }
        $userenabled = $false
    }

    Write-Log "Processing update for $username"
    if ($passive)
    {
        Write-Log "  in passive mode"
    }

    # Verify the user in AD
    try {
        $aduser = Get-ADUser $username -properties memberOf,uidNumber
        $userexists = $true
    }
    catch {
        if ($_.CategoryInfo.Category -eq "ObjectNotFound")
        {
            $userexists = $false
        }
        else
        {
            # Unrecognized AD error. We can't continue
            $global:LastError = "Error fetching $username from AD. Failing: $_"
            Write-Log $LastError
            return 0
        }
    }

    if (-not $userexists -and -not $userenabled)
    {
        # Nothing to do
        Write-Log "User $username is disabled/destroyed and doesn't exist in AD. Nothing to do"
        return 1
    }

    $create = $false
    $update = $false

    # Handle account disable
    # TODO: Revisit how accounts get disabled. For now: 
    #  - prevent logins by disabling account
    #  All other changes (HideInGal, Mailbox removal) are handled by Exchange handler
    # 
    if ($userexists -and -not $userenabled)
    {
        if ($aduser.Enabled -eq $false)
        {
            # Account already disabled. Nothing to do
            Write-Log "User $username is disabled/destroyed and is disabled in AD. Nothing to do"
            return 1
        }
        try {
            if ($passive)
            {
                Write-Log "$($aduser.samAccountName) | Disable-ADAccount"
            }
            else
            {
                Write-Log "Disabling $username"
                $junk = $aduser | Disable-ADAccount
            }
        }
        catch {
            # Unrecognized AD error. We can't continue
            $global:LastError = "Error disabling $username from AD. Failing: $_"
            Write-Log $LastError
            return 0
        }
        return 1
    }

    ## Calculate all of the user attributes we're going to set.
    #
    # Password: 
    # Cipher password has had two different XML tags - cipherpw and cipherPassword, so check 'em both
    $cipherpw = $xmlmsg.synclogin.login.cipherPassword
    if ($cipherpw -eq $null)
    {
        $cipherpw = $xmlmsg.synclogin.login.cipherpw
    }
    if ($cipherpw -and $cipherpw -ne "")
    {
        $newpassword = $crypt.DecryptStringENC($cipherpw)
        if ($newpassword.length -lt 6)
        {
            # This should never happen, so don't let it.
            Write-Log "Warning: $username's new password decrypts to less than 6 chars: $newpassword. Will not set"
            return 0
        }
        if ($global:testusers -contains $username)
        {
            Write-Log " TestUser: password decrypted to $newpassword"
        }
    }
    else
    {
        # Generate a random 24 char password using mixed-case alphanumerics
        $newpassword = -join (((48..57)+(65..90)+(97..122)) * 80 |Get-Random -Count 24 |%{[char]$_})
        if ($passive)
        {
            Write-Log "Passive mode: Random generated password for ${username}: $newpassword"
        }
    }
    $pwcred =  ConvertTo-SecureString "$newpassword" -AsPlainText -Force

    ### Create New User ###
    # If the user doesn't exist yet, create them.
    if ($userenabled -and -not $userexists)
    {
        # Create a new user
        try {
            if ($passive)
            {
                Write-Log "Passive:mode New-ADUser $username -AccountPassword <redacted> -Path $global:UsersOU -SamAccountName $username  -UserPrincipalName $scopedusername"
                Write-Log "   -PasswordNeverExpires $true -ProfilePath '\\%profileserver%\%profileshare%\%username%'"
            }
            else
            {
                Write-Log "Creating new AD account for $username"
                $junk = New-ADUser $username -AccountPassword $pwcred -Path $global:UsersOU -SamAccountName $username -UserPrincipalName $scopedusername `
                             -PasswordNeverExpires $true -ProfilePath '\\%profileserver%\%profileshare%\%username%'
                # If we got here, the account got created without error. Enable it
                $aduser = Get-ADUser $username -properties memberOf,uidNumber
                $aduser | Set-ADUser -Enabled $true
            }
            $create = $true
        } catch {
            # Unrecognized AD error. We can't continue
            $global:LastError = "Error creating $username in AD. Failing: $_"
            Write-Log $LastError
            return 0
        }
        # Fall through to user update to handle additional attributes
    }

    ### Update Existing (and newly created) Users ###
    # Displayname, firstname, lastname
    #
    # Grab roles from Amaint, as that'll determine whether the user is anonymized or not
    $roles = @($xmlmsg.synclogin.person.roles.ChildNodes.InnerText)
    if ($roles -contains "staff" -or $roles -contains "faculty" -or $roles -contains "other" -or [int]$xmlmsg.synclogin.person.sfuVisibility -gt 4)
    {
        $DisplayName = $xmlmsg.synclogin.login.gcos
        $surname = $xmlmsg.synclogin.person.surname
        $firstname = $xmlmsg.synclogin.person.preferredName
    }
    else 
    {
        ### If we want to support true anonymization of students, this is where it'll happen
        # For now, just use their computing ID as their name
        $DisplayName = $username
        $surname = $username
        $firstname = $username
    }

    # Any fields below this point could be blank, but Set-ADuser doesn't let us specify a blank value
    # for an attribute - we have to use the "-clear attributename,[attributename]" syntax.
    # So we will build an array of changes (Replaces and Clears) and pass the entire array to Set-ADUser
    # Google "Powershell splatting" for details
    # Taken from example here: https://social.technet.microsoft.com/Forums/en-US/db692bda-1939-4d04-924c-295cdff1aaa6/setaduser-multiple-attributes?forum=winserverpowershell
    $userprops = @{
        DisplayName = $DisplayName
        Surname = $surname
        GivenName = $firstname
        Replace = @{}
        Clear = @()
    }

    $props = @{}
    # First populate a hashtable with a set of values. Once done, we'll iterate over the
    # values and put the blank ones in the "Clear" pile and the defined ones in the "Replace" pile
    #
    # Fields that only employee/role accounts should have set
    if ($roles -contains "staff" -or $roles -contains "faculty" -or $roles -contains "other")
    {
        $phone = $xmlmsg.synclogin.person.phones.phone
        if ($phone.Count -gt 1)
        {
            # Need support for multi-valued phone number
            # Alan R's script took the *last* phone number and stuffed it into the Office Phone, so
            # we'll do the same. The rest will go into the otherTelephone attribute, which is multi-valued
            $hasphones =  [System.Collections.ArrayList]$phone
            $phone = $hasphones[-1]
            $hasphones.RemoveAt($hasphones.IndexOf($phone))
        }
        $props += @{telephoneNumber = $phone; otherTelephone = $hasphones}
        $props += @{
            url        = $xmlmsg.synclogin.person.url
            Department = $xmlmsg.synclogin.person.galDeptName
            EmployeeID = $xmlmsg.synclogin.person.sfuid
            title      = $xmlmsg.synclogin.person.title
        }
    }
    else
    {
        $props += @{
            telephoneNumber = $null
            otherTelephone = $null
            url = $null
            Department = $null
            EmployeeID = $null
            title = $null
        }
    }

    # "Services For Unix" attributes
    $posixUid = [int]$xmlmsg.synclogin.login.posixUid
    if ($posixUid -eq -1 -and $aduser.uidNumber -gt 0)
    {
        $userprops.Clear += @('UID','loginShell','msSFU30Name','msSFU30NisDomain','unixUSERPassword','gidNumber','unixHomeDirectory','UIDNumber')
    }
    if ($posixUid -gt 0 -and -not ($aduser.uidNumber -gt 0))
    {
        $props += @{
            UID=$username
            loginShell="/bin/csh"
            msSFU30Name=$username
            msSFU30NisDomain="ad"
            unixUSERPassword="ABCD!efgh12345`$67890"
            gidNumber=8088
            unixHomeDirectory="/home/$username"
            UIDNumber=$posixUid
        }
    }

    # Sort attributes into Replace and Clear piles
    $props.Keys | ForEach {
        if ($props.$_)
        {
            $userprops.Replace += @{$_ = $props.$_}
        }
        else
        {
            $userprops.Clear += @($_)
        }
    }
    # If either pile is empty, remove it
    if ($userprops['Replace'].count -eq 0)
    {
        $userprops.Remove('Replace')
    }
    if ($userprops['Clear'].count -eq 0)
    {
        $userprops.Remove('Clear')
    }
 
 
    # Finally, apply all the changes to the user
    try {
        if ($passive)
        {
            Write-Log "Passivemode: Set-ADUser $username $($userprops | convertto-json)"
        }
        else
        {
            if (-not $create)
            {
                Write-Log "Setting user password"
                $aduser | Set-ADAccountPassword -NewPassword $pwcred -Reset
                # If we wanted to, we could also unlock an AD account when we process an update, in case a previous bad
                # password resulted in it getting locked:
                # $aduser | Unlock-ADAccount
            }
            Write-Log "Setting attributes for $username - $($userprops | convertto-json) "
            $aduser | Set-ADUser @userprops
        }
    } catch {
        # Unrecognized AD error. We can't continue
        $global:LastError = "Error updating Unix attributes for $username in AD. Failing: $_"
        Write-Log $LastError
        return 0
    }

    # Update the special "SFU is-<role>" groups, and the is_lightweight group
    try {
        if ($xmlmsg.syncLogin.login.isLightweight -eq "true")
        {
            if ($aduser.memberOf -NotContains "cn=is-lightweight,$global:GroupsOU")
            {
                # Add to the is-lightweight group
                if ($passive)
                {
                    Write-Log "PassiveMode: Add-ADGroupMember -Identity 'cn=is-lightweight,$global:GroupsOU' -Members $username "
                }
                else
                {
                    Write-Log "Adding $username to is-lightweight Group"
                    Add-ADGroupMember -Identity "cn=is-lightweight,$global:GroupsOU" -Members $username   
                }
            }
        }
        elseif ($aduser.memberOf -Contains "cn=is-lightweight,$global:GroupsOU")
        {
            # Account no longer lightweight - remove from lightweight group
            if ($passive)
            {
                Write-Log "PassiveMode: Remove-ADGroupMember -Identity 'cn=is-lightweight,$global:GroupsOU' -Members $username "
            }
            else
            {
                Write-Log "Removing $username from is-lightweight Group"
                Remove-ADGroupMember -Identity "cn=is-lightweight,$global:GroupsOU" -Members $username -Confirm:$false
            }
        }
        
        ForEach ($role in $global:grouproles)
        {
            # Do removes
            if ($aduser.memberOf -contains "CN=SFU is-$role,$global:GroupsOU" -and $roles -NotContains $role)
            {
                if ($passive)
                {
                    Write-Log "PassiveMode: Remove-ADGroupMember -Identity 'cn=SFU is-$role,$global:GroupsOU' -Members $username "
                }
                else
                {
                    Write-Log "Removing $username from group is-$role"
                    Remove-ADGroupMember -Identity "CN=SFU is-$role,$global:GroupsOU" -Members $username -Confirm:$False
                }
            }
            # Do adds
            if ($roles -Contains $role -and $aduser.memberOf -NotContains "CN=SFU is-$role,$global:GroupsOU")
            {
                if ($passive)
                {
                    Write-Log "PassiveMode: Add-ADGroupMember -Identity 'cn=SFU is-$role,$global:GroupsOU' -Members $username "
                }
                else
                {
                    Write-Log "Adding $username to group is-$role"
                    Add-ADGroupMember -Identity "CN=SFU is-$role,$global:GroupsOU" -Members $username
                }
            }
        }
    } catch {
        # Unrecognized AD error. We can't continue
        $global:LastError = "Error updating group memberships for $username in AD. Failing: $_"
        Write-Log $LastError
        return 0
    }

    # We're not going to worry about regular AD group updates here. It would slow us down, and Grouper
    # handles it anyway.
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

## Blowfish Cipher setup
$crypt.CryptAlgorithm = "blowfish2"
$crypt.CipherMode = "cbc"
# An initialization vector is required if using CBC or CFB modes.
# Must match what Amaint used
$ivHex = "`$KJh#(}q"
$crypt.SetEncodedIV($ivHex,"ascii")

if ($global:cipherkey.length -ne 56)
{
    Write-Log("Error: Cipherkey is wrong length. Got $($global:cipherkey.length) bytes. Expected 56. Check $cipherfile")
    exit 1
}
$crypt.KeyLength = 448
$crypt.SetEncodedKey($crypt.Encode($global:cipherkey,"hex"),"hex")

# We need "space" padding mode. See https://www.chilkatsoft.com/refdoc/csCrypt2Ref.html#prop37
$crypt.PaddingScheme = 4

# EncodingMode specifies the encoding of the output for
# encryption, and the input for decryption.
# It may be "hex", "url", "base64", or "quoted-printable".
$crypt.EncodingMode = "base64"

## Done Blowfish setup

$AMQSession = New-ActiveMQSession -Uri $ActiveMQServer -User $Username -Password $Password -ClientAcknowledge

$Target = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$queueName")
$RetryTarget = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$retryQueueName")

# Create a consumer with the target
$Consumer =  $AMQSession.CreateConsumer($Target)
$RetryConsumer = $AMQSession.CreateConsumer($RetryTarget)

# Wait for a message. For now, we'll wait a really short time and 
# if no message arrives, sleep before trying again. That way we can add more logic
# inside our loop later if we want to (e.g. checking multiple queues for messages).
# We sleep for 1 second, so each trip through the loop takes about a second

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
        Remove-ActiveMQSession $AMQSession
        exit 0
    }
}


