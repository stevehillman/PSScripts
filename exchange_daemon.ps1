# This daemon will open up a port and wait for commands
# from a remote client. It allows a non-Windows client
# to issue commands to an Exchange server to
# - get mailbox properties
# - create mailboxes
# - change properties on mailboxes

$ExchangeServer = "http://its-exsv1-tst.exchtest.sfu.ca"
$ListenPort = 2016
$Testing=1
$LogFile = "C:\Users\Hillman\exchange_daemon.log"

# The token we require from the client to verify auth. Simple string compare
$Token = Get-Content "C:\Users\Hillman\exchange_daemon_token.txt" -totalcount 1

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
        write-error "Error connecting to Exchange Server: "
        write-error $_.Exception
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
          do
          {
            $line = $Reader.ReadLine()

            Add-Content $Logfile "Processing command $line"
            # Process command
            if ($line -Match "^$Token getusers")
            {
                $Resp =  get-mailbox | select-object -property PrimarySmtpAddress,`
                                                               RecipientTypeDetails,`
                                                               DisplayName,`
                                                               SamAccountName | ConvertTo-Json
                
            }
            elseif ($line -Match "^$Token getuser ([a-z\-]+)")
            {
                $Resp = Get-Mailbox $Matches[1] | ConvertTo-Json
            }


            elseif ($line -Match "^quit")
            {
                break
            }
            else
            {
                $Writer.write("Unrecognized command $line`n")
                break
            }


            if ($error.Count)
            {
                    $Writer.write("`"Err: Error executing request: $($error.Exception.ToString()) `"`n")
                    $error.Clear()
            }
            else
            {
                $Writer.write($Resp)
                $Writer.write("`n")
            }

          } while ($True)
        }
        catch
        {
            write-error "Lost connection"
        }

        # Close socket and repeat
        $Connection.Close()
    }
}
catch {
    Write-Error $_
}
Finally {
    Write-Host "Cleaning up.."
    $Listener.Stop()
    Remove-PSSession $ESession
}





