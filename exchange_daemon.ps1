# This daemon will open up a port and wait for commands
# from a remote client. It allows a non-Windows client
# to issue commands to an Exchange server to
# - get mailbox properties
# - create mailboxes
# - change properties on mailboxes

[cmdletbinding()]
param([switch]$Testing)

$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
$ListenPort = 2016
$LogFile = "C:\Users\exch taskman\exchange_daemon.log"
$TokenFile = "C:\Users\exch taskman\exchange_daemon_token.txt"
$OU = "SFUUsers"

if ($Testing -eq $True)
{
    $LogFile = "C:\Users\hillman\exchange_daemon.log"
    $TokenFile = "C:\Users\hillman\exchange_daemon_token.txt"
}

# The token we require from the client to verify auth. Simple string compare
$Token = Get-Content $TokenFile -totalcount 1

# Import dependencies
Import-Module -Name PSAOBRestClient

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
        Add-Content $Logfile "Connection from: $($Connection.Client.RemoteEndPoint)"

        $Stream = $Connection.GetStream()
        $Reader = New-Object System.IO.StreamReader $Stream
        $Writer = New-Object System.IO.StreamWriter $Stream
        $Writer.AutoFlush = $True

        $Writer.write("ok`n")

        try {
         
            $line = $Reader.ReadLine()

            Add-Content $Logfile "Processing command $line"
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

                $upn = $samacct + "@its.sfu.ca"
                
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





