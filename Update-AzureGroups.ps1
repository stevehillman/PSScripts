$me = $env:username
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\azuregroups.log"

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:RestToken = $settings.RestToken
    $global:GroupsOU = $settings.AzureGroupsOU
    $global:CompositeAttribute = $settings.AzureGroupsCompositeAttribute
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
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
  return ,$outArray

}


# Import dependencies
Import-Module -Name PSAOBRestClient

load-settings($SettingsFile)

# Get all AD Groups in the target OU
$Groups = GET-ADGroup -Filter '*' -Searchbase $GroupsOU -ResultSetSize $null -Properties $CompositeAttribute -ErrorAction Stop

ForEach ($Group in $Groups) 
{
    Write-Log "Processing Group $($Group.name)"
    $mlmembers = @{}
    $mlexcludes = @{}
    $mlmandatories = @{}
    $members = @{}
    $admembers = @{}
    $HasMandatory = $false
    try {
        if ($Group.$CompositeAttribute -eq "" -or $Group.$CompositeAttribute -eq $null)
        {
            Write-Log "Warning: Skipping $($Group.name). It has no Composite Defintion. Check its $CompositeAttribute attribute"
            continue
        }


        # Gather the memberships from all of the component lists
        # The following are supported
        # - 'listname' or '+listname'   = include members from this list
        # - '-listname'                 = exclude members from this list
        # - '!listname'                 = members MUST be in this list to be included. Multiple '!listname's are ORd together
        # - '!listname1+listname2[+..]' = members MUST be in ALL of the specified lists to be included

        $lists = $Group.$CompositeAttribute.Split(',')
        ForEach ($list in $lists)
        {
            if ($list -match '^\-')
            {
                # Exclude list
                $list1 = $list -replace '^\-',''
                $tempexcludes = Get-AOBRestMaillistMembers -Maillist $list1 -AuthToken $RestToken
                ForEach ($u in $tempexcludes)
                {
                    if ($u -Match "@")
                    {
                        # Skip non-local list members
                        Continue
                    }
                    $mlexcludes[$u] = 1
                }
            }
            elseif ($list -match '^!' )
            {
                $HasMandatory = $true
                $list1 = $list -replace '^!',''
                if ($list1 -match '\+')
                {
                    # Handle ANDed groups. A user must be in ALL of the '+'d groups to be added to the Mandatory list
                    $MandatoryLists = $list1.split('\+')
                    $counter = 0
                    $AndedMembers = @{}
                    $TempAndedMembers = [System.Collections.ArrayList]@()
                    $MinSize = 99999
                    $MinSizeArray = 0

                    # grab the membership of each list that is to be ANDed together
                    ForEach ($l in $MandatoryLists)
                    {
                        $tempmem = Get-AOBRestMaillistMembers -Maillist $l -AuthToken $RestToken
                        if ($tempmem.Count -eq 0)
                        {
                            Write-Log "WARNING: ANDed list $l is empty. Ignoring it"
                            continue
                        }
                        if ($tempmem.Count -lt $MinSize)
                        {
                            $MinSize = $tempmem.Count
                            $MinSizeArray = $counter
                        }
                        $junk = $TempAndedMembers.add($tempmem)
                        $counter = $counter + 1
                    }

                    # Loop through all the members of the *smallest* list
                    ForEach ($u in $TempAndedMembers[$MinSizeArray])
                    {
                        $found = $true
                        # Check the memberships of all of the other lists to see if the user is present
                        ForEach ($i in 0..($counter-1))
                        {
                            if ($i -eq $MinSizeArray)
                            {
                                continue
                            }
                            if ($TempAndedMembers[$i] -notcontains $u)
                            {
                                $found = $false
                                break;
                            }
                        }
                        if ($found)
                        {
                            $mlmandatories[$u] = 1
                        }
                    }
                }
                else
                {
                    # ORd mandatory list
                    $tempmandatory = Get-AOBRestMaillistMembers -Maillist $list1 -AuthToken $RestToken
                    ForEach ($u in $tempmandatory)
                    {
                        if ($u -Match "@")
                        {
                            # Skip non-local list members
                            Continue
                        }
                        $mlmandatories[$u] = 1
                    }
                }
            }
            else 
            {
                # Regular include list
                $list1 = $list -replace '^\+',''
                $tempincludes = Get-AOBRestMaillistMembers -Maillist $list1 -AuthToken $RestToken
                ForEach ($u in $tempincludes)
                {
                    if ($u -Match "@")
                    {
                        # Skip non-local list members
                        Continue
                    }
                    $mlmembers[$u] = 1
                }
            }
        } #ForEach lists

        # We have all of our lists. Calculate final membership
        if ($HasMandatory -and $mlmandatories.count -lt $mlmembers.count)
        {
            # If we have an ANDed list and that list is the smaller, loop through it instead
            # to save cycles

            # We use "PSBase.Keys" JUST IN CASE there's a user named "keys"
            ForEach ($u in $mlmandatories.PSBase.Keys)
            {
                if ($mlmembers[$u] -eq 1 -and $mlexcludes[$u] -ne 1)
                {
                    $members["CN=$u,OU=SFUUsers,DC=ad,DC=sfu,DC=ca"] = 1
                }
            }
        }
        else
        {
            ForEach ($u in $mlmembers.PSBase.Keys)
            {
                if ($HasMandatory -and $mlmandatories[$u] -ne 1)
                {
                    # User isn't in the mandatory maillist(s)
                    continue
                }
                if ($mlexcludes[$u] -eq 1)
                {
                    # User is in the excludes list
                    continue
                }
                $members["CN=$u,OU=SFUUsers,DC=ad,DC=sfu,DC=ca"] = 1
            }
        }

        # Ok, we have our target list of users. Re-fetch the AD Group, retrieving its membership
        $ADGroup = Get-ADGroup $Group.DistinguishedName -Properties members -ErrorAction Stop

        $adds = [System.Collections.ArrayList]@()
        $removes = [System.Collections.ArrayList]@()

        ForEach ($u in $ADGroup.members)
        {
            $admembers[$u] = 1
            if ($members[$u] -ne 1)
            {
                $u1 = $u -replace ',OU=.*','' -replace 'CN=',''
                $junk = $removes.add($u1)
            }
        }

        ForEach ($u in $members.Keys)
        {
            if ($admembers[$u] -ne 1)
            {
                $u1 = $u -replace ',OU=.*','' -replace 'CN=',''
                $junk = $adds.add($u1)
            }
        }

        if ($removes.Count -gt 0 -or $adds.Count -gt 0)
        {
            Write-Log "Processing $($adds.Count) Adds and $($removes.Count) Removes for $($Group.name)"
            if ($removes.Count -gt 0)
            {
                $n = [System.Math]::Ceiling( ($removes.Count / 1000) )
                $Chunks = Split-Array -inArray $removes -parts $n
                Write-Log "Adding $($removes.Count) members in $n chunks of max 1000"
                
                foreach ($Chunk in $Chunks)
                {
                    Remove-ADGroupMember -Identity $Group.DistinguishedName -Members $Chunk -Confirm:$False -ErrorAction Stop 
                }
            }
            if ($adds.Count -gt 0)
            {
                $n = [System.Math]::Ceiling( ($adds.Count / 1000) )
                $Chunks = Split-Array -inArray $adds -parts $n
                Write-Log "Adding $($adds.Count) members in $n chunks of max 1000"
                
                foreach ($Chunk in $Chunks)
                {
                    Add-ADGroupMember -Identity $Group.DistinguishedName -Members $Chunk  -ErrorAction Stop 
                }
            }
            Write-Log "Done processing group $($Group.name)"
        }
        else
        {
            Write-Log "Group $($Group.name) is up to date"
        }
    } catch {
        Write-Log "Error processing update for $Group : $_"
        Continue
    }
}