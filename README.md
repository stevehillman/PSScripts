# PSScripts
A collection of Powershell scripts for managing Exchange

import-allusers.ps1
-------------------

Import existing users from AD into Exchange (create a mailbox).
This script is fairly SFU-specific but easily customized. You
must specify either a maillist or "all" to import all users.
If you specify a maillist, only members of that list will be
imported (assuming they exist in AD). Imported users are set
to be hidden from the GAL and have "_not_migrated" appended
to their email address to ensure they're not visible until
they've been properly migrated from the legacy email system


exchange_daemon.ps1
-------------------

A daemon (designed to be run as a Windows Service) that listens
for TCP connections from other systems and executes certain
commands in the Exchange environment. Primarily used to assist
with migrating users into Exchange, driven from commands from
linux-based mail servers. May also be used for monitoring,
requesting queue lengths, etc.

send-statsd.ps1
---------------

A general purpose script to poll Windows Performance Counters
and send them to a StatsD server. Uses a JSON settings file
to determine which performance counters to monitor and process.
An example settings file is included.

