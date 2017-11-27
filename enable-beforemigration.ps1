# Enable-beforemigration
#
# Make a user's mailbox visible to the Cloud Migrator prior to the start of migration

$userid = '[!user-importname]'
$userid = $userid+"@sfu.ca"

$mb = Get-Mailbox $userid -ErrorAction Stop
if ($mb.PrimarySmtpAddress -Match "\+sfu_connect")
{
	# User needs to be renamed prior to migration
	$emailAddresses = @("SMTP:$userid","smtp:${userid}+sfu_connect")
	Set-Mailbox $userid -EmailAddresses $emailAddresses -ErrorAction Stop
}