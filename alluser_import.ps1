# Import all users in OU=SFUUsers into Exchange. This is meant as a one-time import
# but for each user, checks whether they have a mailbox before creating one

$me = $env:username

# Configurables
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\User-Import.log"

$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
$OU = "OU=SFUUsers,DC=Exchtest,DC=sfu,DC=ca"

$me = $env:username

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:RestToken = $settings.RestToken
    $global:UsersOU = $settings.UsersOU
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
        import-pssession $ESession
}
catch {
        write-host "Error connecting to Exchange Server: "
        write-host $_.Exception
        exit
}


# Fetch all users

$users = GET-ADUser -Filter '*' -Searchbase $UsersOU

foreach ($u in $users)
{
    # Fetch user info from REST
    # Are they lightweight or inactive? If so, 'continue': no need to create
    $amuser = Get-AOBRestUser 

    # For now, we'll consider this a one-time one-way sync of AD users into Exchange,
    # but this script could be modified to be rerunnable, modifying existing Exchange
    # accounts for users whose status has changed - e.g. for disabled/lightweight accounts
    # that DO exist in Exchange, disable them

    $create=$false
	try {
        $mb = Get-Mailbox $u.SamAccountName
    }
    catch {
        # User doesn't have a mailbox yet. Create one
        $create=$true
    }
    if ($create)
    {
        
        try {
            Enable-Mailbox -Identity $u.SamAccountName -Name $u.SamAccountName
            Set-Mailbox -Identity $u.SamAccountName -HiddenFromAddressListsEnabled $true -PrimarySmtpAddress $username+"_not_migrated"
            Write-Log "Created mailbox for $($u.SamAccountName)"
        }
        catch {
            Write-Log "Failed to create mailbox for $($u.SamAccountName). $_"
        }
    }
}