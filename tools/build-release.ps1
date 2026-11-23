# Builds the release file into dist/:
#   dist/teams-calendar.icuewidget   (the iCUE widget package)
# The v1.x helper Setup zip (setup/, helper-ics/) is no longer built; it stays on the v1.5.0 release.
# Requires the iCUE Widget CLI: npm install -g icuewidget-cli
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$dist = Join-Path $root 'dist'
New-Item -ItemType Directory -Force -Path $dist | Out-Null

# --- widget package
Push-Location (Join-Path $root 'widget')
try {
    icuewidget validate TeamsCalendar
    if ($LASTEXITCODE -ne 0) { throw 'Widget validation failed' }
    icuewidget package TeamsCalendar
    if ($LASTEXITCODE -ne 0) { throw 'Widget packaging failed' }
    Copy-Item 'teams-calendar.icuewidget' (Join-Path $dist 'teams-calendar.icuewidget') -Force
} finally { Pop-Location }

# the v1.x setup zip is retired
$oldZip = Join-Path $dist 'TeamsCalendarSetup.zip'
if (Test-Path $oldZip) { Remove-Item $oldZip -Force }

Get-ChildItem $dist | Select-Object Name, Length | Format-Table -AutoSize
