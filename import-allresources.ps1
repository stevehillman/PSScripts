<#
.SYNOPSIS
    Import Resources from a CSV file into Exchange
.DESCRIPTION
    Import all resources in a CSV file into Exchange from AD. This is meant as a one-time import
    but for each resource, it checks whether they have a mailbox before creating one
    It also takes care of adding permissions and can be re-run to apply permissions if, for example,
    the user who was supposed to get the permission didn't exist the first time through
.PARAMETER File
    Name of Resources CSV file
#>

# Force user to provide either a listname or "all"
# If a listname is provided, only members of that list are imported into Exchange
[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$File
    )

$me = $env:username

# Configurables
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\Resource-Import.log"
$MailboxesFile = "C:\Users\$me\mailboxes.json"

# Ensure that Exchange cmdlets throw catchable errors when they fail
$ErrorActionPreference = "Stop"

$GB=1024*1024*1024


function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:RestToken = $settings.RestToken
    $global:UsersOU = $settings.UsersOU
    $global:PassiveMode = ($settings.ImportPassiveMode -eq "true")

}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}


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


# Fetch all users

$users = Get-Content $File | ConvertFrom-CSV     

foreach ($u in $users)
{
    $scopedacct = $u.account
    $acct = $scopedacct -Replace "@sfu.ca"

    if ($acct -NotMatch "^loc-" -and $acct -Notmatch "^equip-")
    {
        # Skip non-resource accts
        Write-Host "Skipping $acct. Not a resource account"
        Continue
    }
    Write-host "Processing $acct"

    # For now, we'll consider this a one-time one-way sync of AD users into Exchange,
    # but this script could be modified to be run on a schedule, modifying existing Exchange
    # accounts for users whose status has changed - e.g. for disabled/lightweight accounts
    # that DO exist in Exchange, disable them

    $create = $false
    try {
        $mb = Get-Mailbox $scopedacct -ErrorAction Stop
    }
    catch {
        $create = $true
    }
    
    if ($create)
    {
        $type = "Room"
        if ($acct -Match "^equip")
        {
            $type = "Equipment"
        }
        if ($PassiveMode)
        {
            Write-Log "PassiveMode: New-mailbox -UserPrincipalName $scopedacct -Displayname $($u.displayname) -$($type)"
        }
        else {
            try {
                if ($type -eq "Room")
                {
                    $junk = New-Mailbox -UserPrincipalName $scopedacct -Displayname $($u.displayname) -Room -ErrorAction Stop
                }
                else {
                    $junk = New-Mailbox -UserPrincipalName $scopedacct -Displayname $($u.displayname) -Equipment -ErrorAction Stop

                }
                Set-Mailbox -Identity $scopedacct -HiddenFromAddressListsEnabled $true `
                            -EmailAddressPolicyEnabled $false `
                            -EmailAddresses "$($acct)+sfu_connect@sfu.ca" `
                            -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update `
                            -ErrorAction Stop
                Set-MailboxMessageConfiguration $userid -IsReplyAllTheDefaultResponse $false -ErrorAction Stop
                Set-CASMailbox $userid -ActiveSyncEnabled $false -ErrorAction Stop
                Write-Log "Created mailbox for $scopedacct"
            }
            catch
            {
                Write-Log "Failed to create mailbox for $($u.SamAccountName). $_"
            }
        }
    }

    # Regardless of whether we just created the account, see if the permissions need updating
   TO DO
}