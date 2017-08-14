# Import all users in OU=SFUUsers into Exchange. This is meant as a one-time import
# but for each user, checks whether they have a mailbox before creating one

# Configurables
$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
$OU = "OU=SFUUsers,DC=Exchtest,DC=sfu,DC=ca"

$me = $env:username

Import-Module -Name PSAOBRestClient

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

$users = GET-ADUser -Filter '*' -Searchbase 'OU=SFUUsers,DC=AD,DC=SFU,DC=CA'

foreach ($u in $users)
{
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
        # Fetch user info from REST
        # Are they lightweight or inactive? If so, no need to create

        # If creating:
        try {
            Enable-Mailbox -Identity $u.SamAccountName -Name $u.SamAccountName
            Set-Mailbox -Identity $u.SamAccountName -HiddenFromAddressListsEnabled $true -PrimarySmtpAddress $username+"_not_migrated"

        }
        catch {
            # What do we do if this fails? Likely just write to a log file
        }
    }
}