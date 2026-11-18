# Builds the two release files into dist/:
#   dist/teams-calendar.icuewidget   (the iCUE widget package)
#   dist/TeamsCalendarSetup.zip      (one-click Setup.cmd + helper for end users)
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

# --- setup zip
$stage = Join-Path ([IO.Path]::GetTempPath()) ('TeamsCalendarSetup-' + [guid]::NewGuid().ToString('N'))
$pkg = Join-Path $stage 'TeamsCalendarSetup'
New-Item -ItemType Directory -Force -Path (Join-Path $pkg 'helper') | Out-Null
foreach ($f in 'Setup.cmd', 'Uninstall.cmd', 'setup.ps1', 'READ ME.txt') {
    # Windows scripts need CRLF line endings
    $text = [IO.File]::ReadAllText((Join-Path $root "setup\$f")) -replace "`r?`n", "`r`n"
    [IO.File]::WriteAllText((Join-Path $pkg $f), $text, (New-Object Text.UTF8Encoding($false)))
}
Copy-Item (Join-Path $root 'helper-ics\ics-calendar-helper.mjs') (Join-Path $pkg 'helper\ics-calendar-helper.mjs')
$zip = Join-Path $dist 'TeamsCalendarSetup.zip'
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $pkg -DestinationPath $zip
Remove-Item $stage -Recurse -Force

Get-ChildItem $dist | Select-Object Name, Length | Format-Table -AutoSize
