# Adds a Startup-folder shortcut so the Teams Calendar Helper v2 (classic Outlook) runs hidden at every Windows sign-in.
# The helper starts classic Outlook minimized when it is not already running.
# Run:   powershell -ExecutionPolicy Bypass -File install-startup.ps1
# Undo:  powershell -ExecutionPolicy Bypass -File install-startup.ps1 -Remove
param([switch]$Remove)

$startup  = [Environment]::GetFolderPath('Startup')
$shortcut = Join-Path $startup 'Teams Calendar Helper (Outlook).lnk'
$v1       = Join-Path $startup 'Teams Calendar Helper.lnk'
$vbs      = Join-Path $PSScriptRoot 'start-hidden.vbs'

if ($Remove) {
    Remove-Item $shortcut -ErrorAction SilentlyContinue
    Write-Host "Removed $shortcut"
    return
}

if (Test-Path $v1) {
    Write-Warning "The v1 (Microsoft Graph) helper is also set to start with Windows. Both use port 38127, so only one can run. Remove v1 with: ..\helper\install-startup.ps1 -Remove"
}

$ws = New-Object -ComObject WScript.Shell
$lnk = $ws.CreateShortcut($shortcut)
$lnk.TargetPath = "$env:WINDIR\System32\wscript.exe"
$lnk.Arguments = "`"$vbs`""
$lnk.WorkingDirectory = $PSScriptRoot
$lnk.Description = 'Serves your classic Outlook calendar to the iCUE Teams Calendar widget'
$lnk.Save()
Write-Host "Created $shortcut"

Start-Process "$env:WINDIR\System32\wscript.exe" -ArgumentList "`"$vbs`""
Write-Host 'Helper started in the background (log: %LOCALAPPDATA%\TeamsCalendarHelper\helper-outlook.log).'
