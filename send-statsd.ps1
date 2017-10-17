# Query Windows performance counters and send the results to a Statsd server
#
# Stats to collect are specified in a JSON settings file
# as a hash of hashes. The key to the outer hash is the 
# common name for the stat (this name will be sent to Statsd)
# The inner hash has the following keys:
#  - Path : Path as defined in Microsoft Performance Counters.
#  - Type : Statsd data type. Defaults to Gauge if omitted
#  - Interval : number of seconds to sample for. Default is 1. The sum of all
#               intervals dictates how often stats will be sent to Statsd
#  - Collapse: "sum|average|zeroaverage". If Path contains a wildcard, specify whether to collapse 
#               all stats returned and whether to sum them or average them, Typically
#               you would sum counters and average response times. Use 'zeroaverage' if you
#               want to include zero values in the average (default is not to, as zeroes may
#               come from counters not actively being updated, throwing off results)
#               If this parameter is left out, each instance will be sent as a separate
#               stat, with the instance name appended to the stat's common name. If any
#               stat is named "_total", it'll be omitted from sum or average
#
# Example:
# "Stats": {
#    "CPU Usage": {
#       "Path": "\\Processor(_Total)\\% Processor Time" 
#   },
#   "HTTP Proxy Reqs": {
#       "Path": "\\MSExchange HttpProxy(*)\\Proxy Requests/Sec",
#       "Collapse": "sum"   
#   },
#   "Disk Free %": {
#       "Path": "\\LogicalDisk(*)\\% Free Space"
#   }
#}

$me = $env:username
$SettingsFile = "C:\Users\$me\statsd-settings.json"
$LogFile = "C:\Users\$me\send-statsd.log"

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:StatsdServer = $settings.StatsdServer
    $global:StatsdPort = $settings.StatsdPort
    $global:Servers = $settings.Servers
    $global:Stats = $settings.Stats
    $global:Namespace = $settings.Namespace
    $global:Debug = ($settings.debug -eq "true")
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

# Set up our UDP Socket ahead of time. It will never change
# Don't wrap these in a 'try'. If server name is invalid, we'll exit immediately.
$UDPclient = new-object System.Net.Sockets.UdpClient; 
$UDPclient.Connect($StatsdServer, $StatsdPort); 

function Write-Statsd($data)
{
    if ($debug)
    {
        Write-Log($data)
        return
    }
    #Encode and send the data
    $encodedData=[System.Text.Encoding]::ASCII.GetBytes($data)
    $bytesSent=$udpclient.Send($encodedData,$encodedData.length)
}


# Main block

# Load settings from json file
load-settings($SettingsFile)

# There are two possible strategies to fetching stats. We can either 
# concat all requested stats into a single get-counter call that will 
# return all stats as a single stat, then parse the results, or
# make a separate get-counter call for each stat. 
# We're going to do the latter because it'll be a lot easier to parse
# the result. But each call to get-counter requires a minimum 1-second 
# to complete, so if we have a *lot* of stats, this won't scale (or it'll
# result in fairly infrequent updates to stats on the Statsd server)

do
{
    $Stats.psobject.Properties | ForEach {
        $statname = $_.Name
        $statpath = $Stats.$statname.Path
        $collapse = $Stats.$statname.Collapse
        $datatype = "g"
        if ($Stats.$statname.Type -eq "c" -or $Stats.$statname.Type -eq "ms")
        {
            $datatype = $Stats.$statname.Type
        }
        $multi = $($statpath -Match "\*" -and $collapse -Match "[a-zA-Z]+")

        $sampleinterval = 1
        if ($Stats.$statname.Interval -gt 1) 
        {
            $sampleinterval = $Stats.$statname.Interval
        }

        $hostdata = @{}
        $hostdatacnt = @{}
        $Servers | ForEach { $hostdatacnt.$_ = 0; $hostdata.$_ = 0 }

        try {
            # Collect the data
            $data = Get-Counter -ComputerName $Servers -sampleinterval $sampleinterval $statpath

            $data.CounterSamples | ForEach {
                # Just in case we can't parse hostname from path, use local hostname as default
                $servername = $env:ComputerName

                # This *should* always match
                if ($_.Path -Match "^\\\\([^\\]+)")
                {
                    $servername = $Matches[1]
                }

                if ($multi)
                {
                    # We're collapsing the values for all instances of a wildcard stat
                    if ($_.InstanceName -eq "_total")
                    {
                        # Skip to the next stat in the ForEach loop
                        return
                    }
                    # Add the value to the total
                    $hostdata.$servername += $_.CookedValue

                    # Add 1 to the number of stats collected for this servername, in case we're averaging
                    if ($_.CookedValue -ne 0 -or $collapse -eq "zeroaverage")
                    {
                        $hostdatacnt.$servername++
                    }
                }
                else 
                {
                    if ($statpath -Match "\*")
                    {
                        $instance = $_.InstanceName -replace "[. (){}/\\:%]","_"
                        $outstring = $Namespace + "." + $servername + "." + $statname + "." + $instance + ":$($_.CookedValue)|$datatype"
                    }
                    else
                    {
                        $outstring = $Namespace + "." + $servername + "." + $statname + ":$($_.CookedValue)|$datatype"

                    }
                    Write-Statsd($outstring)
                }
            }
            if ($multi)
            {
                $hostdata.GetEnumerator() | ForEach {
                    $outdata = $hostdata.$($_.Name)
                    if ($collapse -Match "average" -and $hostdatacnt.$($_.Name) -gt 0)
                    {
                        $$outdata = $outdata / $hostdatacnt.$($_.Name)
                    }
                    Write-Statsd($Namespace + "." + $_.Name + "." + $statname + ":$outdata|$datatype")
                }
            }
        }
        catch {
            Write-Log $_
        }
    }
} while (-Not $debug)
