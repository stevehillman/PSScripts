<#
.SYNOPSIS
    Import AD users into Exchange
.DESCRIPTION
    Import all members of a maillist into Exchange from AD. This is meant as a one-time import
    but for each user, it checks whether they have a mailbox before creating one
.PARAMETER Name
    Name of the maillist to use to determine which users to import. Use "All" to import all users. If importing all users,
    the OU listed in the settings.json is used as the source of users
#>

# Force user to provide either a listname or "all"
# If a listname is provided, only members of that list are imported into Exchange
[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$Name
    )

$me = $env:username

# Configurables
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\User-Import.log"

$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
$OU = "OU=SFUUsers,DC=Exchtest,DC=sfu,DC=ca"

# Ensure that Exchange cmdlets throw catchable errors when they fail
$ErrorActionPreference = "Stop"


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
if ($Name -eq "All")
{
    $users = GET-ADUser -Filter '*' -Searchbase $UsersOU
}
else 
{
    $users = Get-AOBRestMaillistMembers -Maillist $Name -AuthToken $RestToken     
}

foreach ($u in $users)
{
    if ($u -Match "@")
    {
        # Skip non-local list members
        Continue
    }
    if ($Name -ne "All")
    {
        $uad = Get-ADUser $u 
        $u = $uad
        Write-host "Processing $u"
    }
    
    # Fetch user info from REST
    # Are they lightweight or inactive? If so, 'continue': no need to create
    $amuser = Get-AOBRestUser -Username $u.SamAccountName -AuthToken $RestToken
    if ($amuser.isLightweight -eq "true" -or $amuser.status -ne "active")
    {
        Write-Log "Skipping $($u.SamAccountName). Lightweight or Inactive"
        continue
    }

    # For now, we'll consider this a one-time one-way sync of AD users into Exchange,
    # but this script could be modified to be run on a schedule, modifying existing Exchange
    # accounts for users whose status has changed - e.g. for disabled/lightweight accounts
    # that DO exist in Exchange, disable them

    $create = $false
    try {
        $mb = Get-Mailbox $u.SamAccountName
    }
    catch {
        $create = $true
    }
    
    if ($create)
    {
        try {
            Enable-Mailbox -Identity $u.SamAccountName
            Set-Mailbox -Identity $u.SamAccountName -HiddenFromAddressListsEnabled $true `
                        -EmailAddressPolicyEnabled $false `
                        -EmailAddresses "$($u.SamAccountName)_not_migrated@sfu.ca" `
                        -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update
            Set-MailboxMessageConfiguration $u.SamAccountName -IsReplyAllTheDefaultResponse $false
            Write-Log "Created mailbox for $($u.SamAccountName)"
        }
        catch
        {
            Write-Log "Failed to create mailbox for $($u.SamAccountName). $_"
        }
    }
}