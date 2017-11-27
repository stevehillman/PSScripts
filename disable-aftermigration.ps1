# Disable-aftermigration
#
# Make a user's mailbox visible to the Cloud Migrator prior to the start of migration

$userid = '[!user-importname]'
$scopeduserid = $userid+"@sfu.ca"

$mb = Get-Mailbox $scopeduserid -ErrorAction Stop
if ($mb.EmailAddresses -contains "smtp:${userid}+sfu_connect@sfu.ca")
{
	# User needs to be renamed after migration	
	Set-Mailbox $scopeduserid -EmailAddresses "${userid}+sfu_connect@sfu.ca" -ErrorAction Stop
}