<#
=============================================================================================

Name         : Bulk Transfer Meeting Organizer Using PowerShell
Version      : 1.0
Website      : o365reports.com

Script Highlights: 
~~~~~~~~~~~~~~~~~
1. Report-only mode by default to preview eligible meeting transfers. Use -Transfer to perform the transfer.
2. Transfer meetings for a single user or multiple users using a CSV file.
3. Handles recurring meeting series that started before the transfer period but have upcoming occurrences within it.
4. Generates a detailed CSV report with successful, skipped, and failed transfers, including failure reasons.
5. Transfers upcoming meetings for the next 365 days by default. Use -DaysAhead to set a custom time range.
6. Skips meetings starting within the specified lead time using -MinimumLeadHours.
7. Automatically excludes cancelled meetings, attendee-only meetings, and personal appointments unless explicitly included.
8. Validates mailbox types and processes only supported user mailboxes. Shared, room, and group mailboxes are excluded.
9. Supports certificate-based authentication (CBA) and scheduler-friendly execution.
10. Automatically installs the required Exchange Online and Microsoft Graph Calendar modules if not already installed.


For detailed script execution: https://o365reports.com/bulk-transfer-meeting-organizer-using-powershell/

============================================================================================
#>

[CmdletBinding()]
Param(
    [string]$CurrentOrganizer,
    [string]$NewOrganizer,
    [int]$DaysAhead = 365,
    [int]$MinimumLeadHours = 0,
    [string]$CsvPath,
    [Switch]$IncludeAppointmentsWithoutAttendees,
    [string]$TenantId,
    [string]$ClientId,
    [string]$CertificateThumbprint,
    [string]$Organization,
    [Switch]$Transfer,
    [Switch]$Unattended
)

# Module check

Function Connect-ExchangeOnlineSession {
    $ExchangeOnlineModule = Get-Module ExchangeOnlineManagement -ListAvailable
    if ($null -eq $ExchangeOnlineModule) {
        Write-Host "`nImportant: Exchange Online module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $Confirm = Read-Host "Are you sure you want to install Exchange Online module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Exchange Online module..."
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -AllowClobber
            Write-Host "Exchange Online module is installed in the machine successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Exiting.`nNote: Exchange Online module must be available in your system to run the script." -ForegroundColor Red
            Exit
        }
    }

    Import-Module ExchangeOnlineManagement

    Write-Host "`nConnecting to Exchange Online..."
    if ($UseAppOnly) {
        Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbprint `
            -Organization $Organization -ShowBanner:$false -ErrorAction Stop
    }
    else {
        Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }
    Write-Host "Connected to Exchange Online successfully." -ForegroundColor Green
}

Function Connect-MgGraphSession {
    $MsGraphCalendarModule = Get-Module Microsoft.Graph.Calendar -ListAvailable
    if ($null -eq $MsGraphCalendarModule) {
        Write-Host "`nImportant: Microsoft Graph Calendar module is unavailable. It is mandatory to have this module installed in the system to run the script successfully."
        $Confirm = Read-Host "Are you sure you want to install Microsoft Graph Calendar module? [Y] Yes [N] No"
        if ($Confirm -match "[yY]") {
            Write-Host "Installing Microsoft Graph Calendar module..."
            Install-Module Microsoft.Graph.Calendar -Scope CurrentUser -AllowClobber
            Write-Host "Microsoft Graph Calendar module is installed in the machine successfully." -ForegroundColor Magenta
        }
        else {
            Write-Host "Exiting.`nNote: Microsoft Graph Calendar module must be available in your system to discover meetings." -ForegroundColor Red
            Exit
        }
    }

    Import-Module Microsoft.Graph.Calendar

    # Disconnect existing session
    if ($null -ne (Get-MgContext)) {
        Disconnect-MgGraph | Out-Null
    }

    Write-Host "Connecting to Microsoft Graph..."
    if ($UseAppOnly) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId `
            -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
    }
    else {
        # Calendars.Read.Shared covers delegated and shared calendars.
        Connect-MgGraph -Scopes 'Calendars.Read.Shared' -NoWelcome -ErrorAction Stop -WarningAction SilentlyContinue
    }

    if ($null -eq (Get-MgContext)) {
        Write-Host "Failed to connect to Microsoft Graph." -ForegroundColor Red
        Exit
    }
    Write-Host "Connected to Microsoft Graph successfully.`n" -ForegroundColor Green
}

# Transfer logging

Function Write-TransferLog {
    param(
        [string]$FromOrganizer,
        [string]$ToOrganizer,
        [object]$Meeting,
        [string]$Status,
        [string]$Details
    )

    if ($null -eq $Meeting) {
        $Meeting = [PSCustomObject]@{ Subject = "-"; EventId = "-"; MeetingType = "-"; IsOnlineMeeting = "-"; NextOccurrence = "-"; AttendeeCount = "-" }
    }

    $Values = @(
        ($FromOrganizer -replace '"', '""'),
        ($ToOrganizer -replace '"', '""'),
        ([string]$Meeting.Subject -replace '"', '""'),
        $Meeting.MeetingType,
        $Meeting.IsOnlineMeeting,
        $Meeting.NextOccurrence,
        $Meeting.AttendeeCount,
        $Status,
        ($Details -replace '"', '""'),
        ([string]$Meeting.EventId -replace '"', '""')
    )

    if ($null -eq $script:Writer) {
        $script:Writer = New-Object System.IO.StreamWriter -ArgumentList $ExportCSV, $false, ([System.Text.Encoding]::UTF8)
        $script:Writer.AutoFlush = $true
        $script:Writer.WriteLine($script:CsvHeader)
    }
    $script:Writer.WriteLine('"' + ($Values -join '","') + '"')
    $script:Count++
}

# Mailbox address lookup

Function Resolve-OrganizerMailbox {
    param(
        [string]$MailboxIdentity
    )

    # Aliases are resolved to the primary SMTP address.
    $Mailbox = Get-EXOMailbox -Identity $MailboxIdentity -Properties PrimarySmtpAddress, RecipientTypeDetails -ErrorAction Stop

    # Invoke-ChangeMeetingOrganizer supports user mailboxes only.
    if ($Mailbox.RecipientTypeDetails -ne 'UserMailbox') {
        throw "$MailboxIdentity is a $($Mailbox.RecipientTypeDetails). Only user mailboxes are supported."
    }

    return $Mailbox.PrimarySmtpAddress
}

# Meeting discovery

Function Get-OrganizedMeetings {
    param(
        [string]$MailboxAddress,
        [int]$WindowDays
    )

    $WindowStart = (Get-Date).ToUniversalTime()
    $WindowEnd = $WindowStart.AddDays($WindowDays)

    $ViewParams = @{
        UserId        = $MailboxAddress
        StartDateTime = $WindowStart.ToString('yyyy-MM-ddTHH:mm:ssZ')
        EndDateTime   = $WindowEnd.ToString('yyyy-MM-ddTHH:mm:ssZ')
        Property      = @('id', 'subject', 'organizer', 'isOrganizer', 'isCancelled', 'start', 'type', 'seriesMasterId', 'attendees', 'isOnlineMeeting')
        All           = $true
        ErrorAction   = 'Stop'
    }

    # Keyed by series id so a series is transferred once.
    $DiscoveredMeetings = [ordered]@{}

    $Progress = @{ Examined = 0 }

    Get-MgUserCalendarView @ViewParams | ForEach-Object {
        $CalendarEvent = $_
        $Progress.Examined++
        Write-Progress -Activity "Discovering meetings for $MailboxAddress" `
            -Status "Examined $($Progress.Examined) calendar entries"

        # Skip meetings that are already cancelled.
        if ($CalendarEvent.IsCancelled) { return }

        # Keep only meetings the mailbox organizes, not the ones it attends.
        $OrganizerAddress = $CalendarEvent.Organizer.EmailAddress.Address
        if (-not $CalendarEvent.IsOrganizer -and $OrganizerAddress -ne $MailboxAddress) { return }

        # Skip personal appointments with no other attendee.
        $OtherAttendees = @($CalendarEvent.Attendees |
            Where-Object { $_.EmailAddress.Address -and $_.EmailAddress.Address -ne $MailboxAddress })
        if ($OtherAttendees.Count -eq 0 -and -not $IncludeAppointmentsWithoutAttendees) { return }

        # Recurring occurrences transfer through the series id.
        if (-not [string]::IsNullOrWhiteSpace($CalendarEvent.SeriesMasterId)) {
            $MeetingKey = $CalendarEvent.SeriesMasterId
            $MeetingType = "Recurring series"
        }
        else {
            $MeetingKey = $CalendarEvent.Id
            $MeetingType = "Single meeting"
        }

        if ($DiscoveredMeetings.Contains($MeetingKey)) { return }

        $StartDisplay = "-"
        $StartsSoon = $false
        if ($CalendarEvent.Start.DateTime) {
            $StartUtc = [DateTime]$CalendarEvent.Start.DateTime
            $StartDisplay = $StartUtc.ToString("dd-MM-yyyy HH:mm")

            # calendarView returns UTC, so any other zone is labelled.
            if ($CalendarEvent.Start.TimeZone -and $CalendarEvent.Start.TimeZone -ne 'UTC') {
                $StartDisplay += " $($CalendarEvent.Start.TimeZone)"
            }

            # Flag meetings that start inside the lead window.
            if ($MinimumLeadHours -gt 0) {
                $StartsSoon = $StartUtc -lt (Get-Date).ToUniversalTime().AddHours($MinimumLeadHours)
            }
        }

        $SubjectDisplay = if ($CalendarEvent.Subject) { $CalendarEvent.Subject } else { "(No subject)" }

        # Check whether the meeting is an online meeting.
        $OnlineDisplay = if ($CalendarEvent.IsOnlineMeeting) { "Yes" } else { "No" }

        $DiscoveredMeetings[$MeetingKey] = [PSCustomObject]@{
            EventId        = $MeetingKey
            Subject        = $SubjectDisplay
            MeetingType    = $MeetingType
            IsOnlineMeeting = $OnlineDisplay
            NextOccurrence = $StartDisplay
            AttendeeCount  = $OtherAttendees.Count
            StartsSoon     = $StartsSoon
        }
    }

    Write-Progress -Activity "Discovering meetings for $MailboxAddress" -Completed
    return @($DiscoveredMeetings.Values)
}

# Organizer transfer

Function Invoke-OrganizerTransfer {
    param(
        [string]$FromOrganizer,
        [string]$ToOrganizer,
        [object]$Meeting
    )

    # Without -Transfer the meeting is only previewed.
    if (-not $Transfer) {
        Write-TransferLog -FromOrganizer $FromOrganizer -ToOrganizer $ToOrganizer -Meeting $Meeting `
            -Status "Report only" -Details "Pass the switch param `"-Transfer`" to change the organizer"
        return
    }

    $script:Counters.Attempted++

    $TransferParams = @{
        Identity     = $FromOrganizer
        EventId      = $Meeting.EventId
        NewOrganizer = $ToOrganizer
        Confirm      = $false
    }

    $FailureDetail = $null
    try {
        $TransferOutput = Invoke-ChangeMeetingOrganizer @TransferParams 2>&1
        $TransferError = @($TransferOutput | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] })[0]
        if ($TransferError) { $FailureDetail = $TransferError.Exception.Message }
    }
    catch {
        # One meeting failing must not abort the remaining transfers.
        $FailureDetail = $_.Exception.Message
    }

    if ($FailureDetail) {
        $script:Counters.Failed++
        Write-TransferLog -FromOrganizer $FromOrganizer -ToOrganizer $ToOrganizer -Meeting $Meeting `
            -Status "Failed" -Details $FailureDetail
        return
    }

    $script:Counters.Succeeded++
    Write-TransferLog -FromOrganizer $FromOrganizer -ToOrganizer $ToOrganizer -Meeting $Meeting `
        -Status "Succeeded" -Details "Organizer changed"
}

# Meeting transfer for one mailbox

Function Invoke-MailboxMeetingTransfer {
    param(
        [string]$FromOrganizer,
        [string]$ToOrganizer,
        [int]$WindowDays
    )

    # Mailbox validation
    try {
        $CurrentMailbox = Resolve-OrganizerMailbox -MailboxIdentity $FromOrganizer
        $TargetMailbox = Resolve-OrganizerMailbox -MailboxIdentity $ToOrganizer
    }
    catch {
        Write-Host "Cannot process $FromOrganizer to $ToOrganizer : $($_.Exception.Message)" -ForegroundColor Red
        Write-TransferLog -FromOrganizer $FromOrganizer -ToOrganizer $ToOrganizer `
            -Status "Failed" -Details $_.Exception.Message
        return
    }

    Write-Host "`nDiscovering upcoming meetings organized by $CurrentMailbox..." -ForegroundColor Cyan

    # An unreadable calendar must not abort the remaining bulk rows.
    try {
        $Meetings = Get-OrganizedMeetings -MailboxAddress $CurrentMailbox -WindowDays $WindowDays
    }
    catch {
        Write-Host "Calendar discovery failed for $CurrentMailbox : $($_.Exception.Message)" -ForegroundColor Red
        Write-TransferLog -FromOrganizer $CurrentMailbox -ToOrganizer $TargetMailbox `
            -Status "Failed" -Details "Calendar discovery failed: $($_.Exception.Message)"
        return
    }

    if ($Meetings.Count -eq 0) {
        Write-Host "No upcoming organized meetings found for $CurrentMailbox in the next $WindowDays day(s)." -ForegroundColor Yellow
        Write-TransferLog -FromOrganizer $CurrentMailbox -ToOrganizer $TargetMailbox `
            -Status "No meetings found" -Details "No upcoming organized meetings in the next $WindowDays day(s)"
        return
    }

    $script:Counters.Discovered += $Meetings.Count
    Write-Host "Found $($Meetings.Count) meeting(s) to transfer to $TargetMailbox." -ForegroundColor Cyan

    $MeetingIndex = 0
    foreach ($Meeting in $Meetings) {
        $MeetingIndex++
        Write-Progress -Activity "Changing meeting organizer for $CurrentMailbox" `
            -Status "$($Meeting.Subject) (meeting $MeetingIndex of $($Meetings.Count))" `
            -PercentComplete (($MeetingIndex / $Meetings.Count) * 100)

        # A meeting inside the lead window is left for a later run.
        if ($Meeting.StartsSoon) {
            $script:Counters.Skipped++
            Write-TransferLog -FromOrganizer $CurrentMailbox -ToOrganizer $TargetMailbox -Meeting $Meeting `
                -Status "Skipped" -Details "Starts within the $MinimumLeadHours hour lead time"
            continue
        }

        Invoke-OrganizerTransfer -FromOrganizer $CurrentMailbox -ToOrganizer $TargetMailbox -Meeting $Meeting

        # Throttling limits are undocumented, so transfers are paced.
        if ($Transfer) { Start-Sleep -Seconds 1 }
    }

    Write-Progress -Activity "Changing meeting organizer for $CurrentMailbox" -Completed
}

# Report setup

$Location = Get-Location
$ExportCSV = Join-Path $Location "MeetingOrganizerTransfer_$((Get-Date -Format 'yyyy-MMM-dd-ddd_HH-mm-ss').ToString()).csv"
$script:Count = 0
$script:Writer = $null
$script:CsvHeader = '"Current Organizer","New Organizer","Meeting Subject","Meeting Type","Online Meeting","Meeting Start Time (UTC)","Attendees","Status","Details","Event Id"'
$script:Counters = @{ Discovered = 0; Attempted = 0; Succeeded = 0; Skipped = 0; Failed = 0 }

# Mailbox selection

if ($DaysAhead -le 0) {
    Write-Host "-DaysAhead must be a positive integer." -ForegroundColor Red
    return
}

# Prompt for the mailboxes when they are missing.
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    while ([string]::IsNullOrWhiteSpace($CurrentOrganizer)) {
        Write-Host "`nEnter the mailbox that currently organizes the meetings." -ForegroundColor Magenta
        $CurrentOrganizer = (Read-Host "Current organizer").Trim()
    }

    while ([string]::IsNullOrWhiteSpace($NewOrganizer)) {
        Write-Host "Enter the mailbox that takes over the meetings." -ForegroundColor Magenta
        $NewOrganizer = (Read-Host "New organizer").Trim()
    }
}

if (-not [string]::IsNullOrWhiteSpace($CsvPath) -and -not (Test-Path -Path $CsvPath)) {
    Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
    return
}

# App-only needs all four application parameters.
$AppOnlyParameters = @($TenantId, $ClientId, $CertificateThumbprint, $Organization) |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
$UseAppOnly = $AppOnlyParameters.Count -eq 4

if ($AppOnlyParameters.Count -gt 0 -and -not $UseAppOnly) {
    Write-Host "App-only authentication requires -TenantId, -ClientId, -CertificateThumbprint, and -Organization together." -ForegroundColor Red
    return
}

$BulkRows = @()
if (-not [string]::IsNullOrWhiteSpace($CsvPath)) {
    $BulkRows = @(Import-Csv -Path $CsvPath)
    if ($BulkRows.Count -eq 0) {
        Write-Host "The CSV file has no data rows." -ForegroundColor Yellow
        return
    }
}

if (-not $Transfer) {
    Write-Host "`nRunning in report-only mode. No calendar is modified. Pass the switch param `"-Transfer`" to change the organizer." -ForegroundColor Cyan
}

# The transfer notifies external attendees, so confirm once upfront.
if ($Transfer -and -not $Unattended) {
    Write-Host "`nRun without -Transfer to preview what would change without modifying any calendar.`n" -ForegroundColor DarkGray
    $TransferConfirm = Read-Host "Type 'TRANSFER' to proceed"
    if ($TransferConfirm -ne 'TRANSFER') {
        Write-Host "Operation cancelled." -ForegroundColor Yellow
        return
    }
}

# Connect to Exchange Online and Microsoft Graph

Connect-ExchangeOnlineSession
Connect-MgGraphSession

# Process transfers

try {
    if ([string]::IsNullOrWhiteSpace($CsvPath)) {
        Invoke-MailboxMeetingTransfer -FromOrganizer $CurrentOrganizer -ToOrganizer $NewOrganizer `
            -WindowDays $DaysAhead
    }
    else {
        Write-Host "Processing $($BulkRows.Count) transfer request(s)..." -ForegroundColor Cyan

        $RowIndex = 0
        foreach ($Row in $BulkRows) {
            $RowIndex++
            $RowCurrentOrganizer = ([string]$Row.CurrentOrganizer).Trim()
            $RowNewOrganizer = ([string]$Row.NewOrganizer).Trim()

            Write-Progress -Activity "Changing meeting organizer" `
                -Status "$RowCurrentOrganizer (request $RowIndex of $($BulkRows.Count))" `
                -PercentComplete (($RowIndex / $BulkRows.Count) * 100)

            Invoke-MailboxMeetingTransfer -FromOrganizer $RowCurrentOrganizer -ToOrganizer $RowNewOrganizer `
                -WindowDays $DaysAhead
        }
    }
}
finally {
    Write-Progress -Activity "Changing meeting organizer" -Completed
    if ($null -ne $script:Writer) {
        $script:Writer.Close()
        $script:Writer.Dispose()
    }
    Disconnect-ExchangeOnline -Confirm:$false
    Disconnect-MgGraph | Out-Null
}

# Output

Write-Host "`n~~ Script prepared by AdminDroid Community ~~`n" -ForegroundColor Green
Write-Host "~~ Check out " -NoNewline -ForegroundColor Green
Write-Host "admindroid.com" -NoNewline -ForegroundColor Yellow
Write-Host " to get access to 3500+ Microsoft 365 reports and 450+ management actions. ~~`n`n" -ForegroundColor Green

if ($script:Counters.Discovered -gt 0) {
    Write-Host "Discovered $($script:Counters.Discovered) upcoming meeting(s) organized by the given mailbox(es)." -ForegroundColor Cyan
}

if ($script:Counters.Attempted -gt 0 -or $script:Counters.Skipped -gt 0) {
    Write-Host "Summary: $($script:Counters.Attempted) attempted | $($script:Counters.Succeeded) succeeded | $($script:Counters.Skipped) skipped | $($script:Counters.Failed) failed" -ForegroundColor Yellow
}

if (Test-Path -Path $ExportCSV) {
    Write-Host "`nThe meeting organizer transfer log is available in: " -NoNewline -ForegroundColor Yellow
    Write-Host $ExportCSV
    if (-not $Unattended) {
        $Prompt = New-Object -ComObject wscript.shell
        $UserInput = $Prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)
        if ($UserInput -eq 6) {
            Invoke-Item $ExportCSV
        }
    }
}
else {
    Write-Host "`nNo transfer requests were processed."
}
