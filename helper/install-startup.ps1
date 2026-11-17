# Adds a shortcut to your Startup folder so the Teams Calendar Helper runs (hidden) every time you sign in to Windows.
# Run:   powershell -ExecutionPolicy Bypass -File install-startup.ps1
# Undo:  powershell -ExecutionPolicy Bypass -File install-startup.ps1 -Remove
param([switch]$Remove)

$startup  = [Environment]::GetFolderPath('Startup')
$shortcut = Join-Path $startup 'Teams Calendar Helper.lnk'
$vbs      = Join-Path $PSScriptRoot 'start-hidden.vbs'

if ($Remove) {
    Remove-Item $shortcut -ErrorAction SilentlyContinue
    Write-Host "Removed $shortcut"
    return
}

$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($shortcut)
$lnk.TargetPath = "$env:WINDIR\System32\wscript.exe"
$lnk.Arguments = "`"$vbs`""
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.Description = 'Serves your Microsoft 365 calendar to the iCUE Teams Calendar widget'
$lnk.Save()
Write-Host "Created $shortcut"

# Start it now as well
Start-Process "$env:WINDIR\System32\wscript.exe" -ArgumentList "`"$vbs`""
Write-Host 'Helper started in the background (log: %LOCALAPPDATA%\TeamsCalendarHelper\helper.log).'
