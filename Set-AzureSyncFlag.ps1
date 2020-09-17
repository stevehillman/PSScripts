$me = $env:username
$SettingsFile = "C:\Users\$me\settings.json"
$LogFile = "C:\Users\$me\azuread.log"
$UsersOU = "OU=SFUUsers,DC=AD,DC=SFU,DC=CA"
$excluded_users = @("steve","svdasilv","geoffreb","dmerlyn")

function load-settings($s_file)
{
    $settings = ConvertFrom-Json ((Get-Content $s_file) -join "")
    $global:RestToken = $settings.RestToken
}

function Write-Log($logmsg)
{
    Add-Content $LogFile "$(date) : $logmsg"
}

# Import dependencies
Import-Module -Name PSAOBRestClient

load-settings($SettingsFile)

# Load info from two maillists

$consentedusers = Get-AOBRestMaillistMembers -Maillist "its-m365-users" -AuthToken $RestToken
$excluded_users = Get-AOBRestMaillistMembers -Maillist "its-m365-excludes" -AuthToken $RestToken


# Get all AD Users in the target OU
$users = GET-ADUser -Filter '*' -Searchbase $UsersOU -ResultSetSize $null

#Iterate over the users
foreach ($u in $users) {
    if ($excluded_users -contains $u.SamAccountName) {
	    Write-Log "Skipping Excluded User $($u.SamAccountName)"
	    continue
    }
    # Fetch user info from REST
    # Are they lightweight or inactive? If so, 'continue': no need to create
    try {
        $amuser = Get-AOBRestUser -Username $u.SamAccountName -AuthToken $RestToken -ErrorAction Stop
    }
    catch {
        Write-Log "Error: Failed to contact RESTServer for $($u.SamAccountName). Skipping. $_"
        Continue
    }

    if ($amuser.isLightweight -eq "true" -or $amuser.status -ne "active")
    {
        Write-Log "Skipping $($u.SamAccountName). Lightweight or Inactive"
        continue
    }
   
    if ($amuser.roles -contains "staff" -or $amuser.roles -contains "faculty" -or $consentedusers -contains $u.SamAccountName)
    {
	    Write-Log "Setting sync flag for $($u.SamAccountName)"
	    $syncflag = "sync"
    }
    else 
    {
        $syncflag = "nosync"
    } 
    # set-aduser $u -replace @{extensionattribute15=$syncflag} 
    # set-aduser $u -replace @{MSExchExtensionCustomAttribute1=$amuser.roles}
                            
}

