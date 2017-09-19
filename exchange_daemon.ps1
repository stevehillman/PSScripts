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
        $Reader = New-Object System.IO.StreamReader $Stream
        $Writer = New-Object System.IO.StreamWriter $Stream
        $Writer.AutoFlush = $True

        $Writer.write("ok`n")

        try {
         
            $line = $Reader.ReadLine()

            Write-Log $Logfile "Processing command $line"
            # Process command
            if ($line -Match "^$Token getusers")
            {
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
            elseif ($line -Match "^$Token getuser ([a-z\-]+)")
            {
                $Resp = Get-Mailbox $Matches[1] | ConvertTo-Json
            }

            elseif ($line -Match "^$Token getqueue")
            {
                $Resp = Get-ExchangeServer | Get-Queue | ConvertTo-Json
            }

            elseif ($line -Match "^$Token new(user|room|equipment) (.+)")
            {
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
            elseif ($line -Match "^$Token enableuser ([a-z\-]+)")
            {
                $username = $Matches[1]
                try 
                {
                    # Fetch user info from REST
                    # Are they lightweight or inactive? If so, 'continue': no need to create
                    $amuser = Get-AOBRestUser -Username $u.SamAccountName -AuthToken $RestToken
                    if ($amuser.isLightweight -eq "true" -or $amuser.status -ne "active")
                    {
                        Write-Log "Skipping $($u.SamAccountName). Lightweight or Inactive"
                        $Resp = "`"ok. Account lightweight or inactive. Skipping enable`""
                    }
                    else 
                    {
                        $create = $false
                        try {
                            $mb = Get-Mailbox $u.SamAccountName
                        }
                        catch {
                            $create = $true
                            # Clear error stack
                            $error.Clear()
                        }

                        # TODO: To calculate this properly, we need the sfuVisibility flag from Amaint
                        $HideInGal = $true
                        if ($amuser.roles -contains "Staff" -or $amuser.roles -contains "Faculty")
                        {
                            $HideInGal = $false
                        }

                        # TODO: need aliases from REST Server. Not yet available
                        # $addresses = $amuser.aliases
                        # $addresses = $addresses | % { $_ + "@sfu.ca"}
                        
                        if ($create)
                        {
                            try {
                                Enable-Mailbox -Identity $u.SamAccountName
                                Set-Mailbox -Identity $u.SamAccountName -HiddenFromAddressListsEnabled $HideInGal `
                                            -PrimarySmtpAddress "$($u.SamAccountName)@sfu.ca" `
                                            -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update
                                            # -EmailAddresses $addresses
                                Set-MailboxMessageConfiguration $u.SamAccountName -IsReplyAllTheDefaultResponse $false
                                Write-Log "Created mailbox for $($u.SamAccountName)"
                                $Resp = "`"ok. Mailbox created`""

                            }
                            catch
                            {
                                Write-Log "Failed to create mailbox for $($u.SamAccountName). $_"
                            }
                        }
                        else
                        # Mailbox exists (this should virtually always be the case) 
                        {
                            try {
                                Set-Mailbox -Identity $u.SamAccountName -HiddenFromAddressListsEnabled $HideInGal `
                                            -PrimarySmtpAddress "$($u.SamAccountName)@sfu.ca" `
                                            -AuditEnabled $true -AuditOwner Create,HardDelete,MailboxLogin,Move,MoveToDeletedItems,SoftDelete,Update
                                            # -EmailAddresses $addresses
                                Set-MailboxMessageConfiguration $u.SamAccountName -IsReplyAllTheDefaultResponse $false
                                Write-Log "Enabled mailbox for $($u.SamAccountName)"
                                $Resp = "`"ok. Mailbox enabled`""
                            }
                            
                        }
                    }

                }
                catch {
                    write-Log $_.toString()
                    $Writer.write("err: Error executing request: $($_.Exception.Message) `n")
                }
            }


            elseif ($line -Match "^quit")
            {
                # break
                $Resp="bye"
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





