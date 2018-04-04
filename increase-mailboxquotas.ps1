<#
.SYNOPSIS
    Scan mailboxes of faculty/staff/role accounts and increase quotas as needed
.DESCRIPTION
    Scan all mailboxes in Exchange that belong to staff/faculty or role accounts and
    increase their quotas if needed. If they are over a certain threshold, their quota
    will be bumped up by a set amount and they'll be emailed.
    All settings for this script are derived from the settings.json file in the
    home directory of the user executing it

#>

$me = $env:username

# Configurables
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\Increase-mailboxquota.log"


# Ensure that Exchange cmdlets throw catchable errors when they fail
$ErrorActionPreference = "Stop"

$GB=1024*1024*1024

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:SmtpServer = $settings.SmtpServer
    $global:RestToken = $settings.RestToken
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail

    # In Passive mode, say what would be done but don't do it.
    $global:PassiveMode = ($settings.QuotaPassiveMode -eq "true")

    # Array of mail lists containing users to process
    $global:mailboxesToProcess = $settings.mailboxesToProcess

    # If mailbox is greater than this percentage of its quota, bump it.
    $global:quotaThresholdPct = $settings.quotaThresholdPct

    # If mailbox is larger than this size in GB and a bump is needed, send user a warning.
    $global:quotaWarnGB = $settings.quotaWarnGB

    $global:Body1Msg = "This message was sent to inform you that the quota for your SFU Mail account has been automatically increased. No further action is required on your part."

$global:Body2Msg = "Quota increases are now automated, and will be triggered when your usage has reached or exceeded $quotaThresholdPct % of your account quota.

To check your quota usage in your SFU Mail account, you may do the following:
- Login to your account at https://mail.sfu.ca
- Click on Options under the 'gear' symbol in the top-right corner of the screen
- Click on General -> My Account
- Your current quota usage and limit will be displayed.

Concerned that this is a phishing attempt? Contact ITS Service Desk at 778-782-4828.

Regards,

SFU IT Services
Strand Hall 1001
8888 University Drive
Burnaby, BC V5A 1S6
https://www.sfu.ca/itservices
"

$global:WarningMsg = "
NOTE: Although your quota has now been increased above $quotaWarnGB GB, it is recommended you try to keep your total mailbox usage below this limit. 
Depending on what email client you use, you may experience issues with such a large mailbox.
"

}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}


Import-Module -Name PSAOBRestClient

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
        # If the Import fails, it's probably just because the cmdlets have already been imported
        import-pssession $ESession -AllowClobber
}
catch {
        write-host "Error connecting to Exchange Server: "
        write-host $_.Exception
        exit
}

$BumpedUsers = @()
$WarnedUsers = @()

ForEach ($list in $mailboxesToProcess)
{
    $users = Get-AOBRestMaillistMembers -Maillist $list -AuthToken $RestToken     

    foreach ($u in $users)
    {
        if ($u -Match "@")
        {
            # Skip non-local list members
            Continue
        }
        try {
            $uad = Get-ADUser $u
        }
        catch {
            Write-Log "Error: $u does not exist in AD. Skipping"
            Continue
        }
        $u = $uad
        
        # Fetch user info from REST
        # Are they lightweight or inactive? If so, 'continue': no need to create
        # Note, this is probably an unnecessary test as we'll just use auto-generated
        # lists for staff/faculty/sponsored accounts, so those should only ever contain
        # valid accounts

        #try {
        #    $amuser = Get-AOBRestUser -Username $u.SamAccountName -AuthToken $RestToken
        #}
        #catch {
        #    Write-Log "Error: Failed to contact RESTServer for $($u.SamAccountName). Skipping. $_"
        #    Continue
        #}

        #if ($amuser.isLightweight -eq "true" -or $amuser.status -ne "active")
        #{
        #    Write-Log "Skipping $($u.SamAccountName). Lightweight or Inactive"
        #    continue
        #}

        $userid = $u.SamAccountName+"@sfu.ca"
        try {
            $mb = Get-Mailbox $userid -ErrorAction Stop
        }
        catch {
            Write-Log "Error: $u has no mailbox. Skipping"
            Continue
        }
        
        # get current quota
        $oldquota = $mb.IssueWarningQuota
        if ($oldquota -match "(\d+) GB")
        {
            $oldquota = $Matches[1]
            $oldquota = $oldquota/1 # Force to Int
        }
        else {
            # If old quota wasn't set, it'll be "Unlimited" which really just means
            # use the Database default. So use the default.
            $oldquota = 5
        }

        # Get current usage
        try {
            $mbstats = Get-MailboxStatistics -Identity $userid -ErrorAction Stop
        }
        catch {
            Write-Log "Error getting mailbox size for $userid : $_"
            Continue
        }

        $size = $mbstats.TotalItemSize
        if ($size -match "(\d+) GB")
        {
            $size = $Matches[1]
            $size = $size/1
        }
        else {
            # less than 1gb, round to 1 for comparison (which will never fail)
            $size = 1
        }

        # See if any change is needed
        if ($size -gt (($oldquota * $quotaThresholdPct)/100))
        {
            $sendwarning = $false
            $newquota = $oldquota + 1
            if ($size -gt 10)
            {
                # Add 2GB if above 10gb usage
                $newquota++
            }
            if ($size -gt 20)
            {
                # Add 3GB if over 20GB usage
                $newquota++
            }
            if ($newquota -gt $quotaWarnGB)
            {
                $sendwarning = $true
            }

            if ($PassiveMode)
            {
                Write-Host "PassiveMode: Set-Mailbox $userid -IssueWarningQuota $($newquota * $GB) -ProhibitSendQuota $(($newquota+1)*$GB) -ProhibitSendReceiveQuota $(($newquota+2)*$GB)"
            }
            else 
            {    
                try {
                    Set-Mailbox $userid -IssueWarningQuota ($newquota * $GB) -ProhibitSendQuota (($newquota+1)*$GB) -ProhibitSendReceiveQuota (($newquota+2)*$GB) -UseDatabaseQuotaDefaults $false -ErrorAction Stop
                }
                catch {
                    Write-Log "Error. Unable to update quota for $($u.SamAccountName). $_"
                    Continue
                }
            }

            $status = "$userid - Usage: $size GB. Old Quota: $oldquota GB. New Quota: $newquota GB"
            $BumpedUsers += $status
            $outboundmsg = $Body1Msg + "`r`n`r`n" + $status + "`r`n`r`n"

            if ($sendwarning)
            {
                $WarnedUsers += $status
                $outboundmsg += $WarningMsg
            }

            $outboundmsg += $Body2Msg

            if ($PassiveMode)
            {
                Write-Host "PassiveMode: Send-MailMessage -From `"postmast@sfu.ca`" -To $userid -Subject `"Your SFU Mail account quota has been automatically increased`" `
                        -SmtpServer $SmtpServer -Body `"$outboundmsg`""
            }
            else {
                Send-MailMessage -From "postmast@sfu.ca" -To $userid -Subject `"Your SFU Mail account quota has been automatically increased`" `
                        -SmtpServer $SmtpServer -Body $outboundmsg
            }

        }
    }
}

if ($BumpedUsers.count -gt 0)
{
    $msgBody = $BumpedUsers -join "`n"
    if ($WarnedUsers.count -gt 0)
    {
        $msgBody += "`nWarned Users:`n"
        $msgBody += $WarnedUsers -join "`n"
    }

    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "Results of Increase-MailboxQuota Script" -SmtpServer $SmtpServer -Body $msgBody
}
