<#
.SYNOPSIS
    Parses the room list from an Excel file and counts how many meetings
    are scheduled on each room's calendar. Configure script accordingly to accommodate excel column naming.

.REQUIRED APP PERMISSION (Microsoft Graph, Application type, admin-consented)
    - Calendars.Read

.EXAMPLE
    .\Get-RoomMeetingCounts.ps1 -ExcelPath .\Users_With_OlDMST.xlsx
    .\Get-RoomMeetingCounts.ps1 -ExcelPath .\Users_With_OlDMST.xlsx -Days 30
#>

[CmdletBinding()]
param(
    [string]$ExcelPath="C:\Users\Aryan.Ganesh\Downloads\Users_With_OlDMST.xlsx",

    # Window to count scheduled meetings in: from now until now + Days
    [int]$Days = 90,

    [string]$AppId      = '<APP-ID>',
    [string]$Thumbprint = '<THUMBPRINT>',
    [string]$TenantId   = '<TENANT-ID>',

    [string]$OutputCsv  = "C:\TEMP\RoomMeetingCounts_$(Get-Date -Format 'yyyyMMdd_HHmm').csv"
)

$ErrorActionPreference = 'Stop'

# TLS 1.2 pin for Windows PowerShell 5.1
if ($PSVersionTable.PSEdition -ne 'Core') {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

# ---------------------------------------------------------------- modules ----
foreach ($mod in 'Microsoft.Graph.Authentication', 'Microsoft.Graph.Calendar', 'ImportExcel') {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module $mod ..." -ForegroundColor Yellow
        Install-Module $mod -Scope CurrentUser -Force
    }
    Import-Module $mod
}

# ------------------------------------------------------------- read rooms ----
if (-not (Test-Path $ExcelPath)) { throw "Excel file not found: $ExcelPath" }
$rooms = Import-Excel -Path $ExcelPath -WorksheetName 'in'
Write-Host ("Loaded {0} rooms from {1}" -f $rooms.Count, $ExcelPath) -ForegroundColor Cyan

# ---------------------------------------------------------------- connect ----
Write-Host "Connecting to Microsoft Graph ..." -ForegroundColor Cyan
Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome

$startIso = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$endIso   = (Get-Date).AddDays($Days).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')

$results = New-Object System.Collections.Generic.List[object]
$i = 0

foreach ($room in $rooms) {
    $i++
    $upn = $room.UserPrincipalName
    Write-Progress -Activity "Counting scheduled meetings (next $Days days)" -Status $upn -PercentComplete (($i / $rooms.Count) * 100)

    try {
        # Retry up to 3x on Graph throttling
        $events = $null
        for ($attempt = 1; $attempt -le 3; $attempt++) {
            try {
                $events = Get-MgUserCalendarView -UserId $upn `
                            -StartDateTime $startIso -EndDateTime $endIso `
                            -All -PageSize 500 `
                            -Property 'subject,start,isCancelled' -ErrorAction Stop
                break
            }
            catch {
                if ($attempt -eq 3 -or $_.Exception.Message -notmatch '429|TooManyRequests|throttl') { throw }
                Start-Sleep -Seconds (15 * $attempt)
            }
        }

        $events = @($events | Where-Object { -not $_.IsCancelled })
        $next   = $events | Sort-Object { [datetime]$_.Start.DateTime } | Select-Object -First 1

        $results.Add([PSCustomObject]@{
            DisplayName       = $room.DisplayName
            UserPrincipalName = $upn
            ScheduledMeetings = $events.Count
            NextMeeting       = if ($next) { [datetime]$next.Start.DateTime } else { $null }
            NextSubject       = if ($next) { $next.Subject } else { $null }
            Note              = ''
        })
    }
    catch {
        $m = $_.Exception.Message
        $note = switch -Wildcard ($m) {
            '*MailboxNotEnabledForRESTAPI*'   { 'NO CLOUD CALENDAR - mailbox is on-prem (hybrid).' }
            '*REST API is not yet supported*' { 'NO CLOUD CALENDAR - mailbox is on-prem (hybrid).' }
            '*ResourceNotFound*'              { 'Calendar/mailbox not found in cloud.' }
            '*ErrorAccessDenied*'             { 'PERMISSIONS - Calendars.Read (Application) missing or blocked by ApplicationAccessPolicy.' }
            '*Authorization_RequestDenied*'   { 'PERMISSIONS - app lacks required Graph permission.' }
            default                           { "ERROR: $m" }
        }
        $results.Add([PSCustomObject]@{
            DisplayName       = $room.DisplayName
            UserPrincipalName = $upn
            ScheduledMeetings = $null
            NextMeeting       = $null
            NextSubject       = $null
            Note              = $note
        })
    }
}
Write-Progress -Activity 'Counting scheduled meetings' -Completed
Disconnect-MgGraph | Out-Null

$results | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

$booked = $results | Where-Object { $_.ScheduledMeetings -gt 0 } | Sort-Object ScheduledMeetings -Descending
$empty  = $results | Where-Object { $_.ScheduledMeetings -eq 0 }
$failed = $results | Where-Object { $null -eq $_.ScheduledMeetings }

Write-Host ""
Write-Host ("=" * 70)
Write-Host ("Rooms checked                    : {0}" -f $results.Count)
Write-Host ("Rooms with meetings scheduled    : {0}" -f $booked.Count) -ForegroundColor Green
Write-Host ("Rooms with zero meetings         : {0}" -f $empty.Count)  -ForegroundColor Yellow
Write-Host ("Rooms unreadable (on-prem/error) : {0}" -f $failed.Count)
Write-Host ("Report: {0}" -f (Resolve-Path $OutputCsv))
Write-Host ("=" * 70)

if ($booked.Count -gt 0) {
    Write-Host "`nMeetings scheduled per room (next $Days days):" -ForegroundColor Green
    $booked | Format-Table DisplayName, ScheduledMeetings, NextMeeting, NextSubject -AutoSize
}
