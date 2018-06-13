# This daemon will open up a port and wait for commands
# from a remote client. It allows a non-Windows client
# to issue commands to an Exchange server to
# - get mailbox properties
# - create mailboxes
# - change properties on mailboxes

[cmdletbinding()]
param([switch]$Testing)

$ListenPort = 2016
$me = $env:username
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\exchange_daemon.log"
$TokenFile = "C:\Users\$me\exchange_daemon_token.txt"
$OU = "SFUUsers"

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ExchangeServer = $settings.ExchangeServer
    $global:RestToken = $settings.RestToken
    $global:ExchangeUsersListPrimary = $settings.ExchangeUsersListPrimary
    $global:ExchangeUsersListSecondary = $settings.ExchangeUsersListSecondary
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:SmtpServer = $settings.SmtpServer
    $global:Domain = $settings.Domain
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

# Ensure that Exchange cmdlets throw a catchable error when they fail
$ErrorActionPreference = "Stop"

# The token we require from the client to verify auth. Simple string compare
$Token = Get-Content $TokenFile -totalcount 1

# Import dependencies
Import-Module -Name PSAOBRestClient

load-settings($SettingsFile)

# Set up our Exchange shell
$e_uri = $ExchangeServer + "/PowerShell/"
try {
        if ($Testing)
        {
            $Cred = Get-Credential
            $ESession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $e_uri  -Authentication Kerberos -Credential $Cred
        }
        else
        {
            $ESession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $e_uri  -Authentication Kerberos
        }
        import-pssession $ESession
}
catch {
        write-host "Error connecting to Exchange Server: "
        write-host $_.Exception
        exit
}

# Set up our TCP listener

$Listener = [System.Net.Sockets.TcpListener]$ListenPort;
$Listener.Start();
try {
    while ($True)
    {
        if (!$Listener.Pending())
        {
            # No connections waiting. Sleep and retry
            Start-Sleep 1
            Continue
        }
        # Clear error stack
        $error.Clear()

        # Start the connection
        $Connection = $Listener.AcceptTcpClient();
        Write-Log "Connection from: $($Connection.Client.RemoteEndPoint)"

        $Stream = $Connection.GetStream()
        # This should set a 5 second read timeout on input
        $Stream.ReadTimeout = 5000
        $Reader = New-Object System.IO.StreamReader $Stream
        $Writer = New-Object System.IO.StreamWriter $Stream
        $Writer.AutoFlush = $True

        $Writer.write("ok`n")

        try {
         
            $line = $Reader.ReadLine()

            Write-Log "Processing command $line"
            # Process command
            if ($line -Match "^$Token getusers")
            {
                # Fetch all users with mailboxes, returning a subset of attributes. Return as a JSON array
                try {
                    $Resp =  get-mailbox | select-object -property PrimarySmtpAddress,`
                                                               RecipientTypeDetails,`
                                                               DisplayName,`
                                                               SamAccountName | ConvertTo-Json
                }
                catch
                {
                    write-Host $_.toString()
                    $Writer.write("err: Error executing request: $($_.Exception.Message) `n")
                }
                
            }
            elseif ($line -Match "^$Token getuser ([a-z0-9\-_@]+)")
            {
                # Fetch a single user mailbox. Returns all attributes as a JSON hash
                $Resp = Get-Mailbox $Matches[1] | ConvertTo-Json
            }

            elseif ($line -Match "^$Token getqueue")
            {
                # Return all Exchange server queues as a JSON blob
                $Resp = Get-ExchangeServer | % { Get-Queue "$($_.Name)\*" } | ConvertTo-Json
            }
            elseif ($line -Match "^$Token getdatabases")
            {
                # Return summary of Exchange databases. Currently only looks at which
                # server each is on.
                $Resp = Get-MailboxDatabase | Select -Property Name,Server,Servers | ConvertTo-Json
            }

            elseif ($line -Match "^$Token new(user|room|equipment) (.+)")
            {
                # Create a new user or resource mailbox. Newuser functionality will be disabled in prod
                # Attributes of new mailbox are passed in as a JSON string
                $type = $Matches[1]
                $json = $Matches[2]
                $userobj = ConvertFrom-Json $json
                # For a new user, we need account name, firstname, lastname, displayname, password
                $samacct = $userobj.username
                $fn = $userobj.firstname
                $sn = $userobj.lastname

                # Sanitize the input
                # Strip domain, if present
                $samacct = $samacct -replace "@.*",""

                $upn = $samacct + "@" + $Domain
                
                try
                {
                    $spass = ConvertTo-SecureString -String $userobj.password -AsPlainText -Force
                    $Resp = "`"ok`""

                    if ($type -eq "user")
                    {
                        new-mailbox -OrganizationalUnit $OU -UserPrincipalName $upn -Name $samacct -Password $spass -FirstName $fn -Lastname $sn -Displayname $userobj.displayname
                    }
                    elseif ($type -eq "room")
                    {
                        # For rooms and equipment, do we want to enable login to the room/equip account or not? Need to do research
                        # new-mailbox -Room -EnableRoomMailboxAccount $true -UserPrincipalName $upn -RoomMailboxPassword $spass -FirstName $fn -Lastname $sn -Displayname $userobj.displayname
                        new-mailbox -OrganizationalUnit $OU -Room -UserPrincipalName $upn -Name $samacct -FirstName $fn -Lastname $sn -Displayname $userobj.displayname
               
                    }
                    elseif ($type -eq "equipment")
                    {
                        # For rooms and equipment, do we want to enable login to the room/equip account or not? Need to do research
                        # new-mailbox -Equipment -EnableRoomMailboxAccount $true -UserPrincipalName $upn -Password $spass -FirstName $fn -Lastname $sn -Displayname $userobj.displayname
                        new-mailbox -OrganizationalUnit $OU -Equipment -UserPrincipalName $upn -Name $samacct -FirstName $fn -Lastname $sn -Displayname $userobj.displayname
                    }
                }
                catch
                {
                    write-Host $_.toString()
                    # $Writer.write("err: Error executing request: $($_.Exception.Message) `n")
                }
            }
            elseif ($line -Match "^$Token enableuser ([a-z0-9\-_]+)")
            {
                # Enable the mailbox of a user. This is "enabling" in the SFU sense, not Exchange sense.
                # The mailbox's aliases are set properly, removing the "_not_migrated" suffix, and the HideInGal flag is 
                # set according to the account's role(s)
                # If the mailbox doesn't yet exist, it'll be created, but this should very rarely happen
                $username = $Matches[1]
                $scopedusername = $username + "@sfu.ca"
                try 
                {
                    if ($username -Match "^loc-" -or $username -Match "^equip-")
                    {
                        $HideInGal = $false
                        $addresses = @($username)
                        $PreferredEmail = $scopedusername
                    }
                    else 
                    {
                        # Verify the user is in AD. This will fail and be caught by the final 'catch' if the user doesn't exist
                        $aduser = Get-ADUser $username
                    

                        # Fetch user info from REST
                        # Are they lightweight or inactive? If so, 'continue': no need to create
                        $amuser = Get-AOBRestUser -Username $username -AuthToken $RestToken
                        if ($amuser.isLightweight -eq "true" -or $amuser.status -ne "active")
                        {
                            Write-Log "Skipping $username . Lightweight or Inactive"
                            throw "Account lightweight or inactive and can't be enabled"
                        }
                        else 
                        {
                            $HideInGal = $true
                            if ($amuser.roles -contains "staff" -or $amuser.roles -contains "faculty" -or ($amuser.roles -contains "other" -and $amuser.visibility -gt 4))
                            {
                                $HideInGal = $false
                            }

                            $PreferredEmail = $amuser.preferredEmail
                            if ($PreferredEmail -Notmatch "@.*sfu.ca")
                            {
                                # For that rare case when a user has specified a non-SFU PreferredEmail address in SFUDS
                                $PreferredEmail = $username + "@sfu.ca"
                            }

                            $addresses = @($username) + $amuser.aliases

                            ## Security check - if PreferredEmail is an @*sfu.ca address, make sure its one of the user's own addresses
                            if ($PreferredEmail -match "@.*sfu.ca" -and $addresses -notcontains ($PreferredEmail -replace "@.*sfu.ca"))
                            {
                                Write-Log "WARNING: $scopedusername Preferred Email address $PreferredEmail is not one of the their aliases. Ignoring"
                                $PreferredEmail = $scopedusername
                            }
                        }
                    }

                    $create = $false
                    try {
                        $mb = Get-Mailbox $scopedusername -ErrorAction Stop
                    }
                    catch {
                        $create = $true
                        # Clear error stack
                        $error.Clear()
                    }

                    # Preferred address comes first in EmailAddresses list
                    $addresses = @($PreferredEmail) + $addresses

                    $ScopedAddresses = @()
                    ForEach ($addr in $addresses) {
                        if ($addr -Notmatch "@")
                        {
                            $Scopedaddr = $addr + "@sfu.ca"
                        }
                        else 
                        {
                            $Scopedaddr = $addr
                        }
                        if ($ScopedAddresses -contains $Scopedaddr)
                        {
                            # eliminate duplicates
                            continue
                        }
                        $ScopedAddresses += $Scopedaddr
                    }
            
                    try 
                    {
                        if ($create)
                        {
                            Enable-Mailbox -Identity $scopedusername -ErrorAction Stop
                            Write-Log "Created mailbox for $($username)"

                        }
                        Set-Mailbox -Identity $scopedusername -HiddenFromAddressListsEnabled $HideInGal `
                                    -EmailAddressPolicyEnabled $false `
                                    -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update `
                                    -EmailAddresses $ScopedAddresses -ErrorAction Stop
                        Set-CASMailbox $scopedusername -ActiveSyncEnabled $true -OWAEnabled $true  -OwaMailboxPolicy "Default" -ErrorAction Stop
                        Write-Log "Enabled mailbox for $username"
                        $Resp = "ok. Mailbox enabled"

                    }
                    catch
                    {
                        Write-Log "Failed to create mailbox for $username . $_"
                    }
                    try 
                    {
                        # If we just created the mailbox, there's a good chance this command will fail. Since it's not critical
                        # ignore it if it does
                        Set-MailboxMessageConfiguration $scopedusername -IsReplyAllTheDefaultResponse $false -ErrorAction Stop
                    }
                    catch 
                    {
                        Write-Log "Caught error trying to set OWA settings for $scopedusername. Ignoring"    
                    }
                }
                catch {
                    write-Log $_.toString()
                }
            }
            elseif ($line -Match "^$Token disableuser ([a-z0-9\-_]+)")
            {
                # disable (make invisible) a user's mailbox. This will normally only ever be used 
                # if there was a problem migrating a user and the 'enableuser' function needs to be backed out
                $username = $Matches[1]
                $scopedusername = $username + "@sfu.ca"
                try 
                {
                    # Make sure the mailbox for this user actually does exist
                    $mb = Get-Mailbox $scopedusername -ErrorAction Stop

                    # Change alias and hide in GAL
                    Set-Mailbox -Identity $scopedusername -HiddenFromAddressListsEnabled $true `
                                -EmailAddressPolicyEnabled $false `
                                -EmailAddresses "SMTP:$($username)+sfu_connect@sfu.ca" -ErrorAction Stop
                    Set-CASMailbox $scopedusername -ActiveSyncEnabled $false -OWAEnabled $false -ErrorAction Stop
            
                    $Resp = "ok. Mailbox disabled"
                }
                catch {
                    write-Log $_.toString()
                }

            }

            elseif ($line -Match "^quit")
            {
                # break
                $Resp="bye"
            }
            elseif ($line -Match "^forcequit")
            {
                exit 0
            }
            else
            {
                $Resp = "Unrecognized command $line"
                # break
            }


            if ($error.Count)
            {
                    $Writer.write("err: Error executing request: $($error.Exception.Message) `n")
                    $error.Clear()
            }
            else
            {
                $Writer.write($Resp)
                $Writer.write("`n")
            }

         
        }
        catch
        {
            write-Host "Lost connection"
        }

        # Close socket and repeat
        $Connection.Close()
    }
}
catch {
    Write-Host $_
}
Finally {
    Write-Host "Cleaning up.."
    $Listener.Stop()
    Remove-PSSession $ESession
}





