# Disable-aftermigration
#
# Hide a user's mailbox after migration completes. Needs to run
# regardless of success or failure

# Set up new PSSession first

# Set this so that errors are thrown from Cmdlets
$ErrorActionPreference = 'Stop'
# Open the connection
$sesh = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri [!ps-url] -credential (New-object -TypeName System.Management.Automation.PSCredential -argumentlist '[!admin-email]',(ConvertTo-SecureString '[!admin-password]' -AsPlainText -Force)) -Authentication Kerberos -AllowRedirection
# Import the session
$imp = Import-PSSession -Session $sesh -AllowClobber

$userid = '[!user-importname]'
$scopeduserid = $userid+"@sfu.ca"

$mb = Get-Mailbox $scopeduserid -ErrorAction Stop
if ($mb.EmailAddresses -contains "smtp:${userid}+sfu_connect@sfu.ca")
{
    # User needs to be renamed after migration  
    Set-Mailbox $scopeduserid -EmailAddresses "${userid}+sfu_connect@sfu.ca" -AuditEnabled $true -ErrorAction Stop
}

# tear down the session when we're done

Remove-PSSession -Session $sesh
