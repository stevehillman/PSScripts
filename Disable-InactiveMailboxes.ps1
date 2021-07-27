# This script iterates through all Exchange mailboxes, looking for
# ones that belong to inactive/lightweight accounts but have not yet
# been cleaned up. 
# The criteria for cleanup are:
#  - mailbox has been renamed to user_disabled@sfu.ca
#  - user has been lightweight or destroyed for at least a year

$me = $env:username
$LogFile = "C:\Users\$me\inactive_mailboxes.log"
$SettingsFile = "C:\Users\$me\settings.json"

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:SmtpServer = $settings.SmtpServer
    $global:PassiveMode = ($settings.PassiveMode -eq "true")
    $global:AmaintBioURL = $settings.AmaintBioURL
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

load-settings($SettingsFile)

# Set up Exchange PS session

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

$report = ""

$destroycount = 0
$failcount = 0

$global:now = Get-Date -Format FileDate

$Mailboxes = Get-Mailbox -ResultSize unlimited | Where {$_.PrimarySmtpAddress -match "_disabled@sfu.ca"}

ForEach ($Mbox in $Mailboxes)
{
    # Grab the UserBio data from Amaint - it's in the same form as an ActiveMQ message
    # Use it to determine the account state, expiry date, etc
    try {
        $URL = $AmaintBioURL + $mbox.SamAccountName
        $UserBioMsg = Invoke-RestMethod -Method "GET" -Uri $URL 
        [xml]$UserBio = $UserBioMsg
        if ($UserBio.syncLogin.login.destroyDate -match "^[0-9]+" -and 
            ($UserBio.syncLogin.login.isLightweight -eq "true" -or $UserBio.syncLogin.login.status -ne "active"))
        {
            $DestroyDate = $UserBio.syncLogin.login.destroyDate -Replace "T.*" -Replace "-"
            if ($DestroyDate -lt ($now - 10000))
            {
                Write-Log "User $($mbox.SamAccountName) went lightweight or destroyed more than a year ago. Disabling mailbox"
                if (-Not $PassiveMode)
                {
                    $Mbox | Disable-Mailbox
                    $report = $report + "Disabled mailbox for $($mbox.SamAccountName)"
                }
                $destroycount++
            }
        }
    }
    catch {
        Write-Log "Caught error processing $($Mbox.Alias): $_.Exception"
        $report = $report + "Error processing $($Mbox.Alias)"
        $failcount++
    }
}

Write-Log "Completed. Disabled $destroycount mailboxes. Encountered $failcount errors"

if ($report -ne "")
{
    $msgSubj = "Results of Disable-InactiveMailboxes Script"
    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject $msgSubj -SmtpServer $SmtpServer -Body $report
}

