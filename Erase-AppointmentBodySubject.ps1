<#
.SYNOPSIS
    Blank the body and subject for all appointments of a specified resource
.DESCRIPTION
    Iterate through all appointments in a Resource Account's primary calendar and
    set the body and subject of each to blank
.PARAMETER Name
    Name of the resource account
.PARAMETER Url
    URL of the Exchange EWS endpoint or user@domain to trigger autodiscover
.PARAMETER Dry
    Print what changes would be made but don't make them
#>

# Force user to provide a resource account
[cmdletbinding()]
param(
    [parameter(Mandatory=$true)][string]$Mailbox,
    [parameter(Mandatory=$true)][string]$Url,
    [parameter(Mandatory=$false)][switch]$Dry,
    [parameter(Mandatory=$false)][switch]$LeaveAttachments,
    [parameter(Mandatory=$false)][switch]$LeaveSubject,
    [parameter(Mandatory=$false)][switch]$LeaveBody,

    )

# Set up the Exchange environment

# This will only run on a machine where the EWS client lib has been installed

Import-Module -Name 'C:\Program Files\Microsoft\Exchange Server\V15\Bin\Microsoft.Exchange.WebServices.dll'

$exchService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService

# Use credentials of the logged-in user
$exchService.UseDefaultCredentials = $true

if ($Url -match "@")
{
    $exchService.AutoDiscoverUrl($Url)
}
else {
    $exchService.Url = $Url
}

# Set up impersonation
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$Mailbox );


# Loop over a years from 2000 to 2020

ForEach ($year in -18..2)
{
    ForEach ($month in 0..11)
    {
        # Set up a calendar search spanning 1 month (approximately)
        # We only span one month at a time to prevent returning too many results.
        # The LoadPropertiesForItems call below will choke if there are too many items
        $CalView = New-Object  Microsoft.Exchange.WebServices.Data.CalendarView($(Get-Date).AddDays(($year*365)+ ($month*31)), $(Get-Date).AddDays(($year*365) + ($month+1)*31))

        # Fetch all appts from the primary calendar in the given year (Note, Resources will never use secondary calendars)
        $appointments = $exchService.FindAppointments("Calendar",$CalView)

        # Fetch the body for the returned appointments
        $exchService.LoadPropertiesForItems($appointments, [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties)

        if ($Dry)
        {
            $appointments | ft Start,End,Subject,Attachments,Body
        }
        else 
        {
            ForEach $item in $appointments
            {
                $changed = $false
                if ($item.HasAttachments && -Not $LeaveAttachments)
                {
                    $item.Attachments.Clear()
                    $changed = $true
                }
                if (-Not $LeaveSubject && $item.Subject.Length > 0)
                {
                    $item.Subject = ""
                    $changed = $true
                }
                if (-Not $LeaveBody && $item.Body.Text.Length > 0)
                {
                    $item.Body.Text = "[Removed for privacy]"
                    $changed = $true
                }
                if ($changed)
                {
                    # Save back to Exchange
                    $item.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AlwaysOverwrite, [Microsoft.Exchange.WebServices.Data.SendInvitationsOrCancellationsMode]::SendToNone)
                }
            }    
        }

    }
}