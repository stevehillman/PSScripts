# Sigh Microsoft. 
#
# Microsoft 365 Groups, which underpin Teams, do not support being populated with
# a security group, which are all that can be synced from on-premises AD. 
# So we must manually sync security groups to M365 Groups. We do that by
# using an attribute on the on-premises group to tell us which M365 Group
# to sync to. Membership is then compared, and changes made. External users
# (guests) in the M365 Group will be ignored

$me = $env:username
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\azurem365groups.log"

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:RestToken = $settings.RestToken
    $global:GroupsOU = $settings.AzureGroupsOU
    $global:AzureGroupsAdmin = $settings.AzureGroupsAdmin
    $global:AzureGroupsAdminPW = $settings.AzureGroupsAdminPW
    $global:CompositeAttribute = $settings.AzureGroupsCompositeAttribute
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}



# Import modules we need
Import-Module AzureAD

load-settings($SettingsFile)

# Set up AzureAD session
# get credentials and login as AAD admin
$Password = $global:AzureGroupsAdminPW | ConvertTo-SecureString -AsPlainText -Force
$UserCredential = New-Object System.Management.Automation.PSCredential -ArgumentList $global:AzureGroupsAdmin,$Password

Connect-AzureAD -Credential $UserCredential

# Get all AD Groups in the target OU
$Groups = GET-ADGroup -Filter '*' -Searchbase $GroupsOU -ResultSetSize $null -Properties $CompositeAttribute -ErrorAction Stop

ForEach ($Group in $Groups) 
{
    $aadmembers = @{}
    $admembers = @{}
    try {
        if ($Group.$CompositeAttribute -eq "" -or $Group.$CompositeAttribute -eq $null)
        {
            Write-Log "Warning: Skipping $($Group.name). It has no Composite Defintion. Check its $CompositeAttribute attribute"
            continue
        }
        if ($Group.$CompositeAttribute -notmatch "TEAM=")
        {
            continue
        }
        Write-Log "Processing Group $($Group.name)"
        $Params = $Group.$CompositeAttribute.Split(";")
        $AzureTeamGroup = $null
        $AzureADGroup = $null
        ForEach ($p in $Params)
        {
            if ($p -match "^TEAM=")
            {
                $AzureTeamGroup = $p -replace "^TEAM=",""
                break
            }
        }
        if ($AzureTeamGroup -eq $null -or $AzureTeamGroup -eq "")
        {
            Write-Log "Warning: Skipping missing or empty TEAM: $($Group.$CompositeAttribute)"
            continue
        }
        Write-Log "  $($Group.name) is linked to Team $AzureTeamGroup"

        $AzureGroup = Get-AzureADGroup -SearchString $AzureTeamGroup -ErrorAction Stop
        if ($AzureGroup -is [array])
        {
            Write-Log "Can not process $($Group.Name). '$AzureTeamGroup' returns more than one M365 Group"
            continue
        } 
        if ($AzureGroup -eq $null)
        {
            Write-Log "Can not process $($Group.Name). '$AzureTeamGroup' not found in Azure"
            continue
        } 
        if ($AzureGroup.SecurityEnable)
        {
            Write-Log "Can not process $($Group.Name). '$AzureTeamGroup' is an Azure Security Group"
        }

        $AzureMembers = Get-AzureADGroupMember -ObjectID $AzureGroup.ObjectID -All $true -ErrorAction Stop

        $ADGroup = Get-AzureADGroup -SearchString $Group.Name  -ErrorAction Stop
        if ($ADGroup -eq $null -or $ADGroup -eq "")
        {
            Write-Log "Something went wrong loading AzureAD Group $($Group.name). Skipping"
            continue
        }

        $ADGroupMembers = Get-AzureADGroupMember -ObjectID $ADGroup.ObjectID -All $true -ErrorAction Stop

        $adds = [System.Collections.ArrayList]@()
        $removes = [System.Collections.ArrayList]@()

        ForEach ($u in $ADGroupMembers)
        {
            $admembers[$u.UserPrincipalName] = $u.ObjectID
        }

        ForEach ($u in $AzureMembers)
        {
            if ($u.UserPrincipalName -notmatch "@sfu.ca")
            {
                continue
            }
            $aadmembers[$u.UserPrincipalName] = $u.ObjectID
            if ($admembers[$u.UserPrincipalName] -eq $null)
            {
                $junk = $removes.add($u.ObjectID)
            }
        }

        ForEach ($u in $admembers.Keys)
        {
            if ($aadmembers[$u] -eq $null)
            {
                $junk = $adds.add($admembers[$u])
            }
        }

        if ($removes.Count -gt 0 -or $adds.Count -gt 0)
        {
            Write-Log "Processing $($adds.Count) Adds and $($removes.Count) Removes for $($AzureGroup.name)"
            if ($removes.Count -gt 0)
            {
                foreach ($u in $removes)
                {
                    Write-Log "Removing $u"
                    Remove-AzureADGroupMember -ObjectID $AzureGroup.ObjectID -MemberID $u  -ErrorAction Stop 
                }
            }
            if ($adds.Count -gt 0)
            {    
                foreach ($u in $adds)
                {
                    Write-Log "Adding $u"
                    Add-AzureADGroupMember -ObjectID $AzureGroup.ObjectID -RefObjectID $u  -ErrorAction Stop 
                }
            }
            Write-Log "Done processing group $($Group.name)"
        }
    } catch {
        Write-Log "Error processing update for $Group : $_"
        Continue
    }
}