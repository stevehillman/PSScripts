# Enable-beforemigration
#
# Make a user's mailbox visible to the Cloud Migrator prior to the start of migration

# Create session first

# Set this so that errors are thrown from Cmdlets
$ErrorActionPreference = 'Stop'
# Open the connection
$sesh = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri [!ps-url] -credential (New-object -TypeName System.Management.Automation.PSCredential -argumentlist '[!admin-email]',(ConvertTo-SecureString '[!admin-password]' -AsPlainText -Force)) -Authentication Kerberos -AllowRedirection
# Import the session
$imp = Import-PSSession -Session $sesh -AllowClobber

$userid = '[!user-importname]'
$scopeduserid = $userid+"@sfu.ca"

$mb = Get-Mailbox $scopeduserid -ErrorAction Stop
if ($mb.PrimarySmtpAddress -Match "\+sfu_connect")
{
    # User needs to be renamed prior to migration
        # Also disable auditing to spare the audit log the verbosity
    $emailAddresses = @("SMTP:${userid}@sfu.ca","smtp:${userid}+sfu_connect@sfu.ca")
    Set-Mailbox $scopeduserid -EmailAddresses $emailAddresses -AuditEnabled $false -ErrorAction Stop
}

# When we're done, tear down the session

Remove-PSSession -Session $sesh
