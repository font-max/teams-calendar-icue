# Teams Calendar - one-click setup for the iCUE widget's calendar helper.
# Run by double-clicking Setup.cmd. Safe to run again at any time (to change the calendar link,
# repair, or update). Uninstall with Uninstall.cmd.
#
# What it does:
#   1. Makes sure Node.js is installed (offers to install it with winget if missing)
#   2. Installs the helper to %LOCALAPPDATA%\TeamsCalendarHelper\app
#   3. Asks for your published Outlook calendar (ICS) link and tests it
#   4. Starts the helper now and at every Windows sign-in (no console window)
#   5. Copies the widget access key to your clipboard
param(
    [string]$IcsUrl,        # skip the prompt and use this link
    [switch]$KeepLink,      # keep an already-saved link without asking
    [switch]$Uninstall,
    [switch]$Unattended     # no "press Enter" pauses
)

$ErrorActionPreference = 'Stop'
try { $Host.UI.RawUI.WindowTitle = 'Teams Calendar Setup' } catch {}

$DataDir    = Join-Path $env:LOCALAPPDATA 'TeamsCalendarHelper'
$AppDir     = Join-Path $DataDir 'app'
$Port       = 38127
$StartupDir = [Environment]::GetFolderPath('Startup')
$Shortcut   = Join-Path $StartupDir 'Teams Calendar Helper (ICS).lnk'
# Anything that launches a Teams Calendar helper (current or older versions, any install location)
$HelperPattern = 'calendar-helper|teams-calendar-widget|teams-calendar-icue|TeamsCalendarHelper'
$GuideUrl   = 'https://github.com/font-max/teams-calendar-icue#quick-setup'
$PublishUrl = 'https://outlook.office.com/mail/options/calendar/SharedCalendars'

function Step($n, $text) { Write-Host ''; Write-Host "[$n/5] $text" -ForegroundColor Cyan }
function Ok($text)   { Write-Host "  OK   $text" -ForegroundColor Green }
function Info($text) { Write-Host "       $text" }
function Warn($text) { Write-Host "  !    $text" -ForegroundColor Yellow }
function Pause-End   { if (-not $Unattended) { Write-Host ''; Read-Host 'Press Enter to close' | Out-Null } }
function Fail($text) {
    # If an earlier install exists, don't leave it stopped (step 2 stops it to replace files).
    $launcher = Join-Path $AppDir 'start-helper.vbs'
    if ((Test-Path $launcher) -and -not (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)) {
        Start-Process "$env:WINDIR\System32\wscript.exe" -ArgumentList ('"' + $launcher + '"')
    }
    Write-Host ''
    Write-Host "  SETUP STOPPED: $text" -ForegroundColor Red
    Write-Host "  Help: $GuideUrl"
    Pause-End
    exit 1
}
function Ask-YesNo($question) {
    if ($Unattended) { return $true }
    $a = Read-Host "       $question [Y/n]"
    return ($a -eq '' -or $a -match '^(y|yes)$')
}

# Stop any running Teams Calendar helper (v1, ICS or Outlook flavour) - they all share port 38127.
function Stop-Helpers {
    Get-CimInstance Win32_Process -Filter "Name='node.exe'" -ErrorAction SilentlyContinue |
        Where-Object { $_.CommandLine -match 'calendar-helper\.mjs' } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    Start-Sleep -Milliseconds 500
}

# Remove every way an older (or duplicate) helper could start with Windows, so only ours runs.
# -KeepCurrent leaves our own startup shortcut in place.
function Remove-OldAutostart([switch]$KeepCurrent) {
    $removed = @()
    $ws = New-Object -ComObject WScript.Shell
    $dirs = @($StartupDir, [Environment]::GetFolderPath('CommonStartup')) | Where-Object { $_ -and (Test-Path $_) }
    foreach ($d in $dirs) {
        foreach ($f in Get-ChildItem -LiteralPath $d -File -ErrorAction SilentlyContinue) {
            if ($KeepCurrent -and $f.FullName -eq $Shortcut) { continue }
            $what = $f.Name
            if ($f.Extension -eq '.lnk') {
                try { $l = $ws.CreateShortcut($f.FullName); $what = "$($f.Name) $($l.TargetPath) $($l.Arguments)" } catch {}
            } elseif ($f.Extension -match '^\.(vbs|cmd|bat|ps1)$') {
                $what = "$($f.Name) " + (Get-Content -LiteralPath $f.FullName -Raw -ErrorAction SilentlyContinue)
            }
            if ($what -match $HelperPattern -or $f.Name -like 'Teams Calendar Helper*') {
                try { Remove-Item -LiteralPath $f.FullName -Force; $removed += "Startup folder: $($f.Name)" } catch { Warn "Could not remove $($f.FullName) (needs admin?)" }
            }
        }
    }
    foreach ($key in 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run') {
        $props = Get-ItemProperty -Path $key -ErrorAction SilentlyContinue
        if (-not $props) { continue }
        foreach ($p in $props.PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }) {
            if ([string]$p.Value -match $HelperPattern) {
                try { Remove-ItemProperty -Path $key -Name $p.Name -ErrorAction Stop; $removed += "Run key: $($p.Name)" } catch { Warn "Could not remove Run entry '$($p.Name)' (needs admin?)" }
            }
        }
    }
    Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
        (($_.Actions | ForEach-Object { "$($_.Execute) $($_.Arguments)" }) -join ' ') -match $HelperPattern
    } | ForEach-Object {
        try { Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false -ErrorAction Stop; $removed += "Scheduled task: $($_.TaskName)" } catch { Warn "Could not remove scheduled task '$($_.TaskName)' (needs admin?)" }
    }
    return $removed
}

Write-Host ''
Write-Host '  Teams Calendar for iCUE - Setup' -ForegroundColor White
Write-Host '  -------------------------------'

# ---------------------------------------------------------------- uninstall
if ($Uninstall) {
    Stop-Helpers
    Remove-OldAutostart | ForEach-Object { Info "Removed $_" }
    if (Test-Path $DataDir) { Remove-Item $DataDir -Recurse -Force }
    Ok 'Helper stopped and removed (including your saved calendar link and access key).'
    Info 'To also stop sharing your calendar: Outlook > Settings > Calendar > Shared calendars > Unpublish.'
    Info 'Remove the widget itself in iCUE.'
    Pause-End
    exit 0
}

# ---------------------------------------------------------------- 1. Node.js
Step 1 'Checking for Node.js'
function Find-Node {
    $cmd = Get-Command node -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    foreach ($p in @("$env:ProgramFiles\nodejs\node.exe", "${env:ProgramFiles(x86)}\nodejs\node.exe", "$env:LOCALAPPDATA\Programs\nodejs\node.exe")) {
        if ($p -and (Test-Path $p)) { return $p }
    }
    return $null
}
function Node-Major($exe) {
    try { return [int](& $exe -p "process.versions.node.split('.')[0]") } catch { return 0 }
}

$node = Find-Node
if ($node -and (Node-Major $node) -lt 18) {
    Warn "Node.js $(& $node --version) is too old (18 or newer is needed)."
    $node = $null
}
if (-not $node) {
    Info 'The helper needs Node.js (free, from nodejs.org).'
    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget -and (Ask-YesNo 'Install Node.js LTS now?')) {
        Info 'Installing Node.js LTS (Windows may ask for permission)...'
        & winget install --id OpenJS.NodeJS.LTS -e --source winget --accept-package-agreements --accept-source-agreements --silent | Out-Host
        $env:Path = [Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [Environment]::GetEnvironmentVariable('Path', 'User')
        $node = Find-Node
    }
    if (-not $node) {
        Start-Process 'https://nodejs.org/en/download'
        Fail 'Node.js is not installed. Install the LTS version from nodejs.org (opened in your browser), then run Setup again.'
    }
}
Ok "Node.js $(& $node --version)"

# ---------------------------------------------------------------- 2. install helper files
Step 2 'Installing the helper'
$src = @((Join-Path $PSScriptRoot 'helper\ics-calendar-helper.mjs'), (Join-Path $PSScriptRoot '..\helper-ics\ics-calendar-helper.mjs')) |
    Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $src) { Fail 'ics-calendar-helper.mjs was not found next to Setup. Extract the whole zip first, then run Setup.cmd from the extracted folder.' }

New-Item -ItemType Directory -Force -Path $AppDir | Out-Null
Stop-Helpers   # the running helper may be using the old copy
$mjs = Join-Path $AppDir 'ics-calendar-helper.mjs'
Copy-Item -LiteralPath $src -Destination $mjs -Force
Unblock-File -LiteralPath $mjs -ErrorAction SilentlyContinue

# hidden launcher with absolute paths (works even before PATH refreshes after a fresh Node install)
$vbs = Join-Path $AppDir 'start-helper.vbs'
$vbsLine = 'CreateObject("WScript.Shell").Run """' + $node + '"" ""' + $mjs + '"" serve", 0, False'
Set-Content -LiteralPath $vbs -Value $vbsLine -Encoding Default
Ok "Installed to $AppDir"

# ---------------------------------------------------------------- 3. calendar link
Step 3 'Connecting your calendar'
$haveLink = Test-Path (Join-Path $DataDir 'ics-url.dat')
$needLink = $true
if ($IcsUrl) {
    $needLink = $true
} elseif ($haveLink -and ($KeepLink -or (Ask-YesNo 'A calendar link is already saved. Keep using it?'))) {
    $needLink = $false
}

while ($true) {
    if ($needLink) {
        $link = $IcsUrl
        if (-not $link) {
            Info 'Opening Outlook calendar sharing settings in your browser...'
            Start-Process $PublishUrl
            Write-Host ''
            Info 'In the page that opens (sign in if asked):'
            Info '  1. Under "Publish a calendar", choose "Calendar" and "Can view all details"'
            Info '  2. Click Publish'
            Info '  3. Copy the ICS link (the one ending in .ics)'
            Info 'Not there? Outlook > Settings > Calendar > Shared calendars. If "Publish a calendar"'
            Info 'is missing, your organization has turned it off - see the setup guide for other options.'
            Write-Host ''
            $link = Read-Host '       Paste the ICS link here (right-click or Ctrl+V) and press Enter'
        }
        $link = $link.Trim().Trim('"').Trim("'")
        if ($link -notmatch '^(https|webcal)://\S+\.ics(\?\S*)?$') {
            if ($IcsUrl -or $Unattended) { Fail 'That is not an ICS link (it should end in .ics).' }
            Warn 'That does not look like an ICS link (it should end in .ics). Try again.'
            continue
        }
        & $node $mjs set-url $link | Out-Null
        if ($LASTEXITCODE -ne 0) { Fail 'Could not save the calendar link.' }
    }

    Info 'Testing the link...'
    $out = & $node $mjs test 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0) {
        $lines = ($out.Trim() -split "`r?`n")
        Ok $lines[0]
        $lines | Select-Object -Skip 1 -First 8 | ForEach-Object { Info $_.Trim() }
        break
    }
    Warn ('The link did not work: ' + $out.Trim())
    if ($IcsUrl -or $Unattended -or -not (Ask-YesNo 'Paste a different link?')) { Fail 'The calendar link could not be read.' }
    $needLink = $true
    $IcsUrl = $null
}

# ---------------------------------------------------------------- 4. start + autostart
Step 4 'Starting the helper'
Stop-Helpers
$busy = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
if ($busy) {
    $p = Get-Process -Id $busy.OwningProcess -ErrorAction SilentlyContinue
    Fail "Port $Port is used by another program ($($p.ProcessName)). Close it and run Setup again."
}
Remove-OldAutostart | ForEach-Object { Info "Removed an older helper from startup ($_)" }
$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($Shortcut)
$lnk.TargetPath = "$env:WINDIR\System32\wscript.exe"
$lnk.Arguments = '"' + $vbs + '"'
$lnk.WorkingDirectory = $AppDir
$lnk.Description = 'Serves your Outlook calendar to the iCUE Teams Calendar widget'
$lnk.Save()
Start-Process "$env:WINDIR\System32\wscript.exe" -ArgumentList ('"' + $vbs + '"')

$up = $false
for ($i = 0; $i -lt 30 -and -not $up; $i++) {
    Start-Sleep -Milliseconds 500
    try { $up = (Invoke-RestMethod -Uri "http://127.0.0.1:$Port/health" -TimeoutSec 2).ok } catch {}
}
if (-not $up) { Fail "The helper did not start. Log: $DataDir\helper-ics.log" }

# Make sure the program answering on the port is OUR installed helper and accepts the key the widget will use.
$key = (& $node $mjs key | Out-String).Trim()
$owner = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1
$ownerCmd = if ($owner) { (Get-CimInstance Win32_Process -Filter "ProcessId=$($owner.OwningProcess)").CommandLine } else { '' }
if ($ownerCmd -notlike "*$mjs*") { Fail "Another program is answering on port $Port instead of the helper: $ownerCmd" }
try {
    Invoke-RestMethod -Uri "http://127.0.0.1:$Port/api/events" -Headers @{ 'X-Calendar-Key' = $key } -TimeoutSec 60 | Out-Null
} catch {
    Fail "The helper did not accept its own access key ($($_.Exception.Message)). Run Setup again; if it repeats, restart the PC first."
}
$left = @(Get-CimInstance Win32_Process -Filter "Name='node.exe'" -ErrorAction SilentlyContinue | Where-Object { $_.CommandLine -match 'calendar-helper\.mjs' -and $_.CommandLine -notlike "*$mjs*" })
if ($left.Count) { Warn "$($left.Count) other Teams Calendar helper process(es) still running; restarting the PC will clear them." }
Ok 'Running now, verified with your access key, and set to start automatically when you sign in to Windows'

# ---------------------------------------------------------------- 5. access key
Step 5 'Your widget access key'
$copied = $false
try { Set-Clipboard -Value $key; $copied = $true } catch {}
Write-Host ''
Write-Host "        $key" -ForegroundColor Yellow
Write-Host ''
if ($copied) { Ok 'Copied to your clipboard' }
Write-Host ''
Write-Host '  LAST STEP, in iCUE:' -ForegroundColor White
Info '  1. Add the Teams Calendar widget to your Xeneon Edge (if you have not yet)'
Info '  2. If iCUE asks to allow network access to localhost:38127, choose Allow'
Info '  3. In the widget settings, paste the key into "Helper Access Key" (Ctrl+V)'
Write-Host ''
Info 'Your meetings appear within a minute. Run Setup again any time to change the calendar link.'
Pause-End
