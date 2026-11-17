# Adds a Startup-folder shortcut so the Teams Calendar Helper (published ICS link) runs hidden at every Windows sign-in.
# Run:   powershell -ExecutionPolicy Bypass -File install-startup.ps1
# Undo:  powershell -ExecutionPolicy Bypass -File install-startup.ps1 -Remove
param([switch]$Remove)

$startup  = [Environment]::GetFolderPath('Startup')
$shortcut = Join-Path $startup 'Teams Calendar Helper (ICS).lnk'
$vbs      = Join-Path $PSScriptRoot 'start-hidden.vbs'

if ($Remove) {
    Remove-Item $shortcut -ErrorAction SilentlyContinue
    Write-Host "Removed $shortcut"
    return
}

foreach ($other in 'Teams Calendar Helper.lnk', 'Teams Calendar Helper (Outlook).lnk') {
    if (Test-Path (Join-Path $startup $other)) {
        Write-Warning "'$other' is also set to start with Windows. All helpers use port 38127, so only one can run. Remove it with that helper's install-startup.ps1 -Remove."
    }
}

$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($shortcut)
$lnk.TargetPath = "$env:WINDIR\System32\wscript.exe"
$lnk.Arguments = "`"$vbs`""
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.Description = 'Serves your published Outlook calendar to the iCUE Teams Calendar widget'
$lnk.Save()
Write-Host "Created $shortcut"

Start-Process "$env:WINDIR\System32\wscript.exe" -ArgumentList "`"$vbs`""
Write-Host 'Helper started in the background (log: %LOCALAPPDATA%\TeamsCalendarHelper\helper-ics.log).'
