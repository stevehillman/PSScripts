# A powershell script to read from any ActiveMq provider
# 
$ErrorActionPreference = "Stop"

Import-Module -Name PSActiveMQClient
Import-Module -Name PSGrouperClient

$me = $env:username
$LogFile = "C:\Users\$me\activemq_grouper_client.log"
$SettingsFile = "C:\Users\$me\settings.json"


## Local private functions ##

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:ActiveMQServer = $settings.ActiveMQServer
    $global:Username = $settings.amqUsername
    $global:Password = $settings.amqPassword
    $global:queueName = $settings.GrouperQueueName
    $global:retryQueueName = $settings.GrouperRetryQueueName
    $global:GrouperUser = $settings.GrouperUser
    $global:GrouperPassword = $settings.GrouperPassword
    $global:GroupsOU = $settings.GroupsOU
    $global:UsersOU = $settings.UsersOU
    $global:RestToken = $settings.RestToken
    $global:MaxRetries = $settings.MaxRetries
    $global:MaxRetryTimer = $settings.MaxRetryTimer
    $global:ErrorsFromEmail = $settings.ErrorsFromEmail
    $global:ErrorsToEmail = $settings.ErrorsToEmail
    $global:MaxNoActivity = $settings.MaxNoActivity
    $global:SmtpServer = $settings.SmtpServer
    $global:PassiveMode = ($settings.GrouperPassiveMode -ne "false")
    
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff') : $logmsg"
}

function Split-array 
{

<#  
  .SYNOPSIS   
    Split an array
  .NOTES
    Version : July 2, 2017 - implemented suggestions from ShadowSHarmon for performance   
  .PARAMETER inArray
   A one dimensional array you want to split
  .EXAMPLE  
   Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -parts 3
  .EXAMPLE  
   Split-array -inArray @(1,2,3,4,5,6,7,8,9,10) -size 3
#> 

  param($inArray,[int]$parts,[int]$size)
  
  if ($parts) {
    $PartSize = [Math]::Ceiling($inArray.count / $parts)
  } 
  if ($size) {
    $PartSize = $size
    $parts = [Math]::Ceiling($inArray.count / $size)
  }

  $outArray = New-Object 'System.Collections.Generic.List[psobject]'

  for ($i=1; $i -le $parts; $i++) {
    $start = (($i-1)*$PartSize)
    $end = (($i)*$PartSize) - 1
    if ($end -ge $inArray.count) {$end = $inArray.count -1}
	$outArray.Add(@($inArray[$start..$end]))
  }
  return $outArray

}

# Scan an array of users to see if they exist in AD. Return an array of those that do.
# Since AD users are never deleted (today, anyway), we can cache successes and skip them next time

$GoodUserLookup = @{}

function Check-ADUsers($UserArray)
{
    $GoodUsers = @{}
    ForEach ($adu in $UserArray)
    {
        if ($GoodUserLookup[$adu] -eq 1)
        {
            Continue
        }
        try {
            $junk = Get-ADUser $adu -Server $pdc -ErrorAction Stop 
        }
        catch {
            continue
        }
        $GoodUsers[$adu] = 1
        $GoodUserLookup[$adu] = 1
    }
    return $GoodUsers.PSBase.Keys
}

function process-message($jsonmsg)
{
    # All Grouper messages are esbEvent messages, so this better never fail
    # esbEvent is actually an array. Grouper only supports one event at a time
    # but could support more in the future
    if ($jsonmsg.esbEvent)
    {
        ForEach ($esbEvent in $jsonmsg.esbEvent)
        {
            $rc = process-grouper-message($esbEvent)
            if ($rc -eq 0)
            {
                return 0
            }
        }
        return 1
    }
    # Add other message types here in the future
    else
    {
        Write-Log "Ignoring msg: Unsupported type"
        return 1
    }
}

# Compare two arrays for differences and return a Custom Object with two properties - an array of OnlyInOne and an array of OnlyInTwo
# To achieve performance, we use the LINQ MS Framework. This is complete overkill for a few hundred members, but for lists with many 
# thousands, it becomes significant. Code was shamelessly stolen from Stack Overflow
#
# https://stackoverflow.com/questions/6368386/comparing-two-arrays-get-the-values-which-are-not-common
#
# See here for more info on using LINQ with Powershell: https://www.red-gate.com/simple-talk/dotnet/net-framework/high-performance-powershell-linq/

# This function has not yet been tested

function compare-arrays($arrayobj1, $arrayobj2)
{
    # First, Cast as Strings, just in case they aren't
    [string[]]$array1 = $arrayobj1
    [string[]]$array2 = $arrayobj2
    $onlyInOne = [string[]][Linq.Enumerable]::Except($array1, $array2)
    $onlyInTwo = [string[]][Linq.Enumerable]::Except($array2, $array1)
    $value = "" | Select-Object -Property OnlyInOne, OnlyInTwo
    $value.OnlyInOne = $onlyInOne
    $value.OnlyInTwo = $onlyInTwo

    return $value
}

# Process a Grouper ActiveMQ message.
# At SFU, we've got custom ActiveMQ code that runs on the AMQ server and "squashes" Grouper event messages
# (which consist of individual add/drop member events) down into a single "AD Update" message. These messages
# are also rate-limited to 1 per minute, to ensure AD doesn't get overloaded. So if 100 users are added to a list,
# only a single AD Update message will be sent.
#
# However, the AMQ code then does 3 things:
# - calls the AmaintRest API to determine all AD groups that have this group as a member
# - ForEach group, fetch all members (also done via AmaintRest), and write the members to a file named for the group name
# - FTP each file to a Windows server (tecnically this step is done using a Camel route, rather than our custom code)
#
# The old Windows DOS scripts would then process each of those files, except that it wasn't, as Alan found it didn't work for him. He
# was relying on GDD instead
#
# As such, this code has nothing to process yet, as we don't want to continue the process of using FTP to move data
#
# In order to simplify the AMQ code, we will strip it down - just send the group names, once per minute, in AMQ messages. 
# The below code then needs to:
# - decode the message (Grouper messages are JSON, not XML, so it's a bit different)
# - Fetch parent groups via GrouperRest
# - Iterate through each resulting group and
#   - if it doesn't have the ADGroup flag set, skip
#   - fetch the flattened membership of the group from Grouper
#   - compare the membership with what is currently in AD
#   - apply changes in chunks

 

function process-grouper-message($esbEvent)
{
    $GroupName = $esbEvent.groupName
    # Fetch this group's details
    $ThisGroup = Get-GrouperGroup -Group $GroupName -Username $global:GrouperUser -Password $global:GrouperPassword
    
    # Add it to the (empty) list of groups to process
    $PotentialGroups = @($ThisGroup)

    # Fetch the parent groups. If there are any, add them
    $resp = Get-GrouperMemberships -Group $GroupName -Username $global:GrouperUser -Password $global:GrouperPassword
    if ($resp.Count)
    {
        $PotentialGroups = $PotentialGroups + $resp
    }

    $UpdateGroups = @()

    # Look at the resulting list of groups and see which ones we care about
    # Grouper group attributes are a bit weird. They're contained in two separate arrays - the attribute names in
    # one, and the attribute values in another, with the same corresponding index value. So see if the attribute
    # name exists, and if it does, use its index to fetch the corresponding value
    ForEach ($PGroup in $PotentialGroups)
    {
        if ($PGroup.detail.attributeNames -contains "sfuIsADGroup" -and 
            $PGroup.detail.attributeValues[[array]::indexof($PGroup.detail.attributeNames,"sfuIsADGroup")] -eq 1)
        {
            $UpdateGroups = $UpdateGroups + @($PGroup.name)
        }    
        else
        {
            # Group is not an AD group. We should test to see whether it exists in AD. If it does, it needs to be deleted
            # TODO: Process group deletes
        }
    }

    # Quick return if there's nothing to do
    if ($UpdateGroups.count -eq 0)
    {
        Write-Log "$($ThisGroup.name) is not an AD group and does not belong to any AD groups. Skipping"
        return 1
    }

    # We have our list of groups to sync. For each group, fetch the flattened memberships from Grouper
    # and the current members from AD and compare. If the group doesn't exist in AD yet, create it.
    ForEach ($Pgroup in $UpdateGroups) {
        # Group names will be in grouper naming format (stem:groupname).
        # In the (near?) future, we will likely need to support multiple stems and
        # update different OUs based on that. E.g.
        # - AzureGroups
        # - Teams
        # - Exchange DLs
        #
        # For now, we'll only process groups in the Maillist stem, so skip any others
        if ($PGroup -notmatch "^maillist:")
        {
            Write-Log "Skipping non-maillist group $PGroup"
            continue
        }

        $ADGroupName =  "CN=" + $PGroup.Substring($PGroup.lastindexof(':')+1) + "," + $global:GroupsOU
        Write-Log "  Processing $ADGroupName"

        # If it's an add or update, test to see whether group exists. Fetch the members while we're at it
        $groupexists = $true
        try {
            $ADGroup = Get-ADGroup $ADGroupName -Properties members -Server $pdc -ErrorAction Stop
        }
        catch {
            if ($_.CategoryInfo.Category -eq "ObjectNotFound")
            {
                $groupexists = $false
            }
            else
            {
                # Unrecognized AD error. We can't continue
                $global:LastError = "Error fetching $ADGroupName in AD. Failing: $_"
                Write-Log $LastError
                return 0
            }
        }

        if (-not $groupexists)
        {
            # Group doesn't exist. Create it
            Write-Log "Group $ADGroupName doesn't exist. Attempting to create it"
            try {
                if ($global:PassiveMode)
                {
                    Write-Log "PassiveMode: New-ADGroup $ADGroupName -Path $global:GroupsOU -Server $pdc"
                    $ADGroup = @{
                        members = @()
                    }
                }
                else
                {
                    New-ADGroup $ADGroupName -Path $global:GroupsOU -GroupCategory Security -GroupScope Global -Server $pdc -ErrorAction Stop
                    Write-Log "Group $ADGroupName created"
                    $ADGroup = Get-ADGroup $ADGroupName -Properties members -Server $pdc -ErrorAction Stop
                }
            }
            catch {
                # Unrecognized AD error. We can't continue
                $global:LastError = "Error creating $ADGroupName in AD. Failing: $_"
                Write-Log $LastError
                return 0
            }
        }
        Write-Log "    Fetched $($ADGroup.members.count) members from AD"

        # Fetch Grouper memberships
        $GrouperGroup = Get-GrouperGroup -Group $PGroup -Members -OnlyUsers -Username $global:GrouperUser -Password $global:GrouperPassword
        Write-Log "    Fetched $($GrouperGroup.members.count) members from Grouper"
        
        # Convert Grouper memberships (userIDs) into AD memberships. For now this just means converting them into Distinguished Names
        # Arrays aren't dynamically resizeable, so use a hash - this has proven fast in other cases
        $GrouperMembers = @{}
        ForEach ($gm in $GrouperGroup.Members)
        {
            $GrouperMembers["CN=$gm,$global:UsersOU"] = 1
        }
        Write-Log "      Converted to ADUsers array" # This can be removed later - it's only here for performance calculation purposes

        # compare-arrays doesn't work if either array is empty, so only use it if both sides have some members
        if ($ADGroup.Members.Count -gt 0 -and $GrouperGroup.Members.Count -gt 0)
        {
            # Now pass both arrays to our fast array comparator to produce two lists of diffs
            # We use "PSBase.Keys" JUST IN CASE there's a user named "keys"
            $AddsDrops = compare-arrays $ADGroup.Members.toLower() $GrouperMembers.PSBase.Keys.toLower()
            $toAdd = $AddsDrops.OnlyInTwo
            $toDrop = $AddsDrops.OnlyInOne
        }
        else
        {
            if ($ADGroup.Members.Count -eq 0)
            {
                # New or empty AD group needs populating
                $toAdd = $GrouperMembers.PSBase.Keys
                $toDrop = @()
            }
            else
            {
                # AD group needs emptying out
                # This should actually be a warning - we may not want to process this, in case the empty group
                # was actually a failure from Grouper
                $toAdd = @()
                $toDrop = $ADGroup.Members
            }
        }

        # If there are Adds or Drops, break them up into chunks with each chunk containing 1000 users
        if ($ToAdd.Count)
        {
            $n = [System.Math]::Ceiling( ($ToAdd.Count / 1000) )
            # Work around a powershell idiosynchracy: if the number of resulting chunks is one,
            # powershell will replace the array of arrays with a single array, screwing up our for loop below.
            # So force it to create an array of arrays
            if ($n -eq 1)
            {
                $Chunks = @("junk")
                $Chunks[0] = $toAdd
            }
            else
            {
                $Chunks = Split-Array -inArray $ToAdd -parts $n
            }
            Write-Log "Adding $($toAdd.Count) members in $n chunks of max 1000"
            if ($toAdd.Count -gt 10)
            {
                $addlog = $toAdd[0..10] -join "`r`n"
            }
            else
            {
                $addlog = $toAdd -join "`r`n"
            }
            Write-Log "  Chunks: $($Chunks.count). Adding users: $addlog"
            foreach ($Chunk in $Chunks)
            {
                $retryadd = 0
                try { 
                    Add-ADGroupMember -Identity $ADGroupName -Members $Chunk -Server $pdc -ErrorAction Stop -WhatIf:$global:PassiveMode
                }
                catch { 
                    if ($_.CategoryInfo.Category -eq "ObjectNotFound" -and $_.CategoryInfo.TargetType -eq "ADPrincipal")
                    {
                        # One or more members doesn't exist in AD
                        $retryadd = 1
                        Write-Log "Error: Chunk has non-existent users"
                    }
                    else
                    {
                        Write-Log "Error adding users to $ADGroupName : $_"
                        return 0
                    }
                }
                if ($retryadd)
                {
                    # Remove the non-existent users from this chunk and try again
                    $checkedChunk = Check-ADUsers $Chunk
                    if ($checkedChunk.count -gt 0)
                    {
                        Write-Log "  Trying again with $($checkedChunk.count) users in Chunk"
                        try { 
                            Add-ADGroupMember -Identity $ADGroupName -Members $checkedChunk -Server $pdc -ErrorAction Stop -WhatIf:$global:PassiveMode
                        }
                        catch { 
                            Write-Log "Error adding users to $ADGroupName : $_"
                            return 0
                        }
                    }
                    else
                    {
                        Write-Log "  After removing invalid users, none left to add"
                    }
                }
            }
        }
        if ($toDrop.Count)
        {
            $n = [System.Math]::Ceiling( ($toDrop.Count / 1000) )
            if ($n -eq 1)
            {
                $Chunks = @("junk")
                $Chunks[0] = $toDrop
            }
            else
            {
                $Chunks = Split-Array -inArray $ToDrop -parts $n
            }
            Write-Log "Removing $($toDrop.Count) members in $n chunks of max 1000"
            if ($toDrop.Count -gt 10)
            {
                $addlog = $toDrop[0..10] -join "`r`n"
            }
            else
            {
                $addlog = $toDrop -join "`r`n"
            }
            Write-Log "  Removing users: $addlog"

            if (-not $global:PassiveMode)
            {
                foreach ($Chunk in $Chunks)
                {
                    $retryrm = 0
                    try { 
                        Remove-ADGroupMember -Identity $ADGroupName -Members $Chunk -Server $pdc -ErrorAction Stop -Confirm:$false
                    }
                    catch { 
                        Write-Log "Error removing users from $ADGroupName : $_"
                        return 0
                    }
                }
            }
        }
    }
    # If we got here, we succeeded
    Write-Log "Group $GroupName successfully processed"
    return 1
}


# Queue a message in the retry queue to retry it later.
# We reformat the XML - wrap it in a "retryMessage" tag and
# add a retry count tag.
function retry-message($m)
{
    $mtmp = $m.Text | ConvertFrom-Json
    $firstRetry = $false
    # Add a retry counter if one isn't already there
    if (! $mtmp.retryCount)
    {
        $mtmp | Add-Member -NotePropertyName retryCount -NotePropertyValue 1
    }
    # Otherwise add one to the retry count
    else
    {
        $count = [int]$mtmp.retryCount
        if ($count -eq 1)
        {
            $firstRetry = $true
        }
        $count++
        $mtmp.retryCount = "$count"
    }

    if ([int]$mtmp.retryCount -gt $MaxRetries)
    {
        Write-Log "FAIL. Max retries exceeded for $($m.Text)"
        return 0
    }

    $outmsg = $mtmp | ConvertTo-Json -Depth 10

    Send-ActiveMQMessage -Queue $retryQueueName -Session $AMQSession -Message $outmsg

    if ($firstRetry)
    {
        return 2
    }
    return 1

}

## end local functions



## main code block

load-settings($SettingsFile)

Write-Log "Starting up"

## Fetch the domain controller we'll use
$pdcobj = Get-ADDomainController -Discover -Service PrimaryDC
$pdc = $pdcobj.HostName[0]
Write-Log "Using $pdc for our domain controller"

$AMQSession = New-ActiveMQSession -Uri $ActiveMQServer -User $Username -Password $Password -ClientAcknowledge

$Target = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$queueName")
$RetryTarget = [Apache.NMS.Util.SessionUtil]::GetDestination($AMQSession, "queue://$retryQueueName")

# Create a consumer with the target
$Consumer =  $AMQSession.CreateConsumer($Target)
$RetryConsumer = $AMQSession.CreateConsumer($RetryTarget)

# Wait for a message. For now, we'll wait a really short time and 
# if no message arrives, sleep before trying again. That way we can add more logic
# inside our loop later if we want to (e.g. checking multiple queues for messages)

$loopcounter=1
$noactivity=0
$retryTimer=10
$retryFailures=0
$msg=""

while(1)
{
    try {
        $isRetry = $false
        $Message = $Consumer.Receive([System.TimeSpan]::FromTicks(10000))
        if (!$Message)
        {
            # No message from the main queue. See if we should check the retry queue
            $loopcounter++
            # Only try the retry queue every x seconds, where x is 10x number of failures in a row
            # This acts as a sequential backoff timer. The retry queue is first tried after 10 seconds
            # then if a message in the retry queue fails again, the next retry time will be 20 seconds,
            # and so on. That way, if there's a failure in underlying infrastructure, the message should
            # still eventually get processed. 
            if ($loopcounter -gt $retryTimer)
            {
                $Message = $RetryConsumer.Receive([System.TimeSpan]::FromTicks(10000))
                # Only if no message was found, reset loop counter so that if there are 
                # multiple messages to be tried, they'll all be tried at once
                if (!$Message) 
                { 
                    $loopcounter = 1
                    # Also reset the number of retry Failures, since there's no msgs left
                    if ($retryFailures)
                    {
                        Write-Log "No messages in retry queue. Clearing retryFailures" 
                        $retryFailures = 0
                        $retryTimer=10
                    }
                }
                # Also reset the loop counter if last retrymsg failed, so that we back off properly.
                elseif ($retryFailures) { $loopcounter = 1 }
            }
            if (!$Message)
            {
                $noactivity++
                if ($noactivity -gt $MaxNoActivity)
                {
                    $noactivity=0
                    Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "No Grouper activity from ActiveMQ for $MaxNoActivity seconds" `
                    -SmtpServer $SmtpServer -Body "Seems a bit fishy."
                }
                Start-Sleep -Seconds 1
                continue
            }

            # Got a message from the Retry queue. Extract the inner message
            $isRetry=$true
            # undef the msg variable before defining it, because retry msgs and regular msgs are slightly different object types
            Remove-Variable msg
            $msg = $Message.Text | ConvertFrom-Json
            Write-Log "Retrying msg `r`n$($Message.Text)"
        }
        else
        {
            # undef the msg variable before defining it, because retry msgs and regular msgs are slightly different object types
            Remove-Variable msg
            $msg = $Message.Text | ConvertFrom-Json
        }

        $noactivity=0

        if (-Not $isRetry) { Write-Log "Processing msg `r`n $($Message.Text)" }
        $rc = process-message($msg)
        Write-Log "RC = $rc"
        if ($rc -gt 0)
        {
            Write-Log "Success"
            $rc = $Message.Acknowledge()
            if ($isRetry) 
            { 
                $retryFailures = 0 
                $retryTimer=10
            }
        }
        else
        {
            if ($isRetry) 
            { 
                $retryFailures++ 
                $retryTimer = (1+$retryFailures) * 10
                if ($retryTimer -gt $MaxRetryTimer) { $retryTimer = $MaxRetryTimer }
                Write-Log "Retry backoff is now $retryTimer seconds"
            }
            Write-Log "Failure. Will Retry"
            $rc = retry-message($Message)

            if ($rc -eq 2)
            {
                # Special case: The first time processing a msg in the retry queue, if it fails, treat it as a success
                # for the purposes of calculating the backoff, so that if there are messages queued behind it, we get
                # through them all quickly.
                $retryFailures = 0
            }
            # Even if retry-message exceeds max retries, we still have to Acknowledge msg to clear it from the queue
            $Message.Acknowledge()
            if ($rc -eq 0)
            {
                Send-MailMessage -From $ErrorsFromEmail -To $ErrorsToEmail -Subject "Failure from Grouper ActiveMQ handler" `
                    -SmtpServer $SmtpServer -Body "Failed to process message $MaxRetries time.`r`nMessage: $($Message.Text). `r`nLast Error: $LastError"
            }
        }

    }
    catch {
        $_
        # Realistically, we want to log errors but try to recover
        # For now we'll just exit and let Windows Scheduler restart us
        write-host "Caught error. Closing sessions"
        Remove-ActiveMQSession $AMQSession
        exit 0
    }
}


