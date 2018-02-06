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
    [parameter(Mandatory=$false)][int]$YearStart = 2000,
    [parameter(Mandatory=$false)][int]$YearEnd = 2025,
    [parameter(Mandatory=$false)][switch]$LeaveAttachments,
    [parameter(Mandatory=$false)][switch]$LeaveSubject,
    [parameter(Mandatory=$false)][switch]$LeaveBody

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
$exchService.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,$Mailbox );


# Loop over a years from 2000 to 2020

ForEach ($week in (($YearStart-2018)*52)..(($YearEnd-2018)*52))
{
    # Set up a calendar search spanning 1 month (approximately)
    # We only span one month at a time to prevent returning too many results.
    # The LoadPropertiesForItems call below will choke if there are too many items
    $CalView = New-Object  Microsoft.Exchange.WebServices.Data.CalendarView($(Get-Date).AddDays($week*7), $(Get-Date).AddDays(($week+1)*7))

    # Fetch all appts from the primary calendar in the given year (Note, Resources will never use secondary calendars)
    $appointments = $exchService.FindAppointments("Calendar",$CalView)

    if ($appointments.TotalCount -eq 0)
    {
        #nothing found. Next.
        Continue
    }

    # Fetch the body for the returned appointments
    $junk = $exchService.LoadPropertiesForItems($appointments, [Microsoft.Exchange.WebServices.Data.PropertySet]::FirstClassProperties)

    if ($Dry)
    {
        $appointments | ft Start,End,Subject,Attachments,Body
    }
    else 
    {
        ForEach ($appt in $appointments)
        {
            $item = $appt
            # We need to check whether this is a recurring event. If it is, we'll never be able to change
            # each occurrence if it has no end date, so fetch the master and change it there. 
            if ($item.AppointmentType -eq "Occurrence")
            {
                # try to fetch the master
                $master =  [Microsoft.Exchange.WebServices.Data.Appointment]::BindToRecurringMaster($exchService, $item.Id)
                if ($master.AppointmentType -eq "RecurringMaster")
                {
                    $item = $master
                }
            }

            $changed = $false
            if ($item.HasAttachments -And -Not $LeaveAttachments)
            {
                $item.Attachments.Clear()
                $changed = $true
            }
            if ((-Not $LeaveSubject) -And $item.Subject.Length -gt 0)
            {
                $item.Subject = ""
                $changed = $true
            }
            if ((-Not $LeaveBody) -And $item.Body.Text.Length -gt 0 -And $item.Body.Text -Notmatch "\[Removed for privacy\]")
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