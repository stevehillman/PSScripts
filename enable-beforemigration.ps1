# Enable-beforemigration
#
# Make a user's mailbox visible to the Cloud Migrator prior to the start of migration

$userid = '[!user-importname]'
$scopeduserid = $userid+"@sfu.ca"

$mb = Get-Mailbox $scopeduserid -ErrorAction Stop
if ($mb.PrimarySmtpAddress -Match "\+sfu_connect")
{
	# User needs to be renamed prior to migration
	$emailAddresses = @("SMTP:${userid}@sfu.ca","smtp:${userid}+sfu_connect@sfu.ca")
	Set-Mailbox $scopeduserid -EmailAddresses $emailAddresses -ErrorAction Stop
}