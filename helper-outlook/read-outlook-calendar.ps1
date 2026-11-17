# Reads today's and tomorrow's appointments from classic Outlook (desktop) via its COM object model
# and writes them to stdout as JSON. Called by outlook-calendar-helper.mjs every couple of minutes.
# Outlook does the Microsoft 365 sign-in (email, password, 2FA) itself; no app registration is involved.
param([int]$Days = 2)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Write-Result($obj) {
    $obj | ConvertTo-Json -Depth 5 -Compress
    exit 0
}

# 1. Is classic Outlook set up (has a mail profile)?
$profiles = Get-ChildItem 'HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles' -ErrorAction SilentlyContinue
if (-not $profiles) {
    Write-Result @{ status = 'needs_login'; message = 'Classic Outlook is not set up yet. Open classic Outlook (outlook.exe) and sign in with your work account.' }
}

# 2. Make sure classic Outlook is running (start it minimized if not), then attach to it.
function Get-Outlook {
    try { return [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') } catch { return $null }
}
$ol = Get-Outlook
if (-not $ol) {
    if (-not (Get-Process OUTLOOK -ErrorAction SilentlyContinue)) {
        try { Start-Process 'outlook.exe' -WindowStyle Minimized } catch {
            Write-Result @{ status = 'error'; message = 'Could not start classic Outlook (outlook.exe).' }
        }
    }
    for ($i = 0; $i -lt 30 -and -not $ol; $i++) { Start-Sleep -Seconds 2; $ol = Get-Outlook }
    if (-not $ol) {
        Write-Result @{ status = 'starting'; message = 'Waiting for classic Outlook to start. If it is asking you to sign in, finish signing in.' }
    }
}

# 3. Read the default calendar, expanding recurring meetings.
try {
    $ns = $ol.GetNamespace('MAPI')
    $cal = $ns.GetDefaultFolder(9)   # olFolderCalendar
    $user = ''
    try { $user = $ns.CurrentUser.Name } catch {}

    $start = (Get-Date).Date
    $end = $start.AddDays($Days)
    $items = $cal.Items
    $items.Sort('[Start]')
    $items.IncludeRecurrences = $true
    $filter = "[Start] < '" + $end.ToString('g') + "' AND [End] > '" + $start.ToString('g') + "'"
    $found = $items.Restrict($filter)

    $teamsProp = 'http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl'
    $linkRe = 'https://(?:teams\.microsoft\.com/l/meetup-join/[^\s<>"]+|teams\.live\.com/meet/[^\s<>"]+|[\w.-]*zoom\.us/[jw]/[^\s<>"]+|meet\.google\.com/[a-z0-9-]+|[\w.-]+\.webex\.com/[^\s<>"]+)'
    $busy = @{ 0 = 'free'; 1 = 'tentative'; 2 = 'busy'; 3 = 'oof'; 4 = 'workingElsewhere' }
    $resp = @{ 0 = 'none'; 1 = 'organizer'; 2 = 'tentativelyAccepted'; 3 = 'accepted'; 4 = 'declined'; 5 = 'notResponded' }

    $events = New-Object System.Collections.ArrayList
    foreach ($it in $found) {
        if ($events.Count -ge 300) { break }
        if ($it.Start -ge $end) { break }

        $join = ''
        try { $join = [string]$it.PropertyAccessor.GetProperty($teamsProp) } catch {}
        if (-not $join) {
            try { $m = [regex]::Match([string]$it.Body, $linkRe); if ($m.Success) { $join = $m.Value } } catch {}
        }

        $private = ($it.Sensitivity -eq 2 -or $it.Sensitivity -eq 3)
        $organizer = ''
        try { $organizer = [string]$it.Organizer } catch {}

        if ($it.AllDayEvent) {
            $s = $it.Start.ToString('yyyy-MM-dd') + 'T00:00:00Z'
            $e = $it.End.ToString('yyyy-MM-dd') + 'T00:00:00Z'
        } else {
            $s = $it.Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            $e = $it.End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        }

        [void]$events.Add([ordered]@{
            key         = [string]$it.GlobalAppointmentID + '|' + $s
            subject     = $(if ($private) { 'Private appointment' } elseif ($it.Subject) { [string]$it.Subject } else { '(No title)' })
            start       = $s
            end         = $e
            isAllDay    = [bool]$it.AllDayEvent
            location    = $(if ($private) { '' } else { [string]$it.Location })
            organizer   = $organizer
            joinUrl     = $join
            isOnline    = [bool]$join
            showAs      = $(if ($busy.ContainsKey([int]$it.BusyStatus)) { $busy[[int]$it.BusyStatus] } else { 'busy' })
            isCancelled = ($it.MeetingStatus -eq 5 -or $it.MeetingStatus -eq 7)
            response    = $(if ($resp.ContainsKey([int]$it.ResponseStatus)) { $resp[[int]$it.ResponseStatus] } else { 'none' })
        })
    }

    Write-Result ([ordered]@{ status = 'ok'; user = @{ name = $user }; events = @($events) })
} catch {
    Write-Result @{ status = 'error'; message = 'Outlook error: ' + $_.Exception.Message }
}
