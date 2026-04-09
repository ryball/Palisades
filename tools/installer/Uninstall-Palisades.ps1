[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [switch]$RemoveUserData
)

$ErrorActionPreference = 'Stop'

$appName = 'DeskHaven'
$installDirs = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\DeskHaven'),
    (Join-Path $env:LOCALAPPDATA 'Programs\Palisades')
)
$startMenuDirs = @(
    (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\DeskHaven'),
    (Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Palisades')
)
$desktopShortcutPaths = @(
    (Join-Path ([Environment]::GetFolderPath('Desktop')) 'DeskHaven.lnk'),
    (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Palisades.lnk')
)
$startupShortcutPaths = @(
    (Join-Path ([Environment]::GetFolderPath('Startup')) 'DeskHaven.lnk'),
    (Join-Path ([Environment]::GetFolderPath('Startup')) 'Palisades.lnk')
)
$registryPaths = @(
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskHaven',
    'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Palisades'
)
$startupRunKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$startupRunValueNames = @('DeskHaven', 'Palisades')
$userDataDir = Join-Path $env:LOCALAPPDATA 'Palisades'

Get-Process DeskHaven, Palisades -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

foreach ($path in ($desktopShortcutPaths + $startupShortcutPaths)) {
    if (Test-Path $path) {
        Remove-Item $path -Force -ErrorAction SilentlyContinue
    }
}

foreach ($startMenuDir in $startMenuDirs) {
    if (Test-Path $startMenuDir) {
        Remove-Item $startMenuDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

foreach ($registryPath in $registryPaths) {
    if (Test-Path $registryPath) {
        Remove-Item $registryPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if (Test-Path $startupRunKeyPath) {
    foreach ($startupRunValueName in $startupRunValueNames) {
        Remove-ItemProperty -Path $startupRunKeyPath -Name $startupRunValueName -ErrorAction SilentlyContinue
    }
}

foreach ($installDir in $installDirs) {
    if (Test-Path $installDir) {
        $cleanupScriptPath = Join-Path $env:TEMP ("DeskHaven-Uninstall-" + [guid]::NewGuid().ToString('N') + '.cmd')
        @"
@echo off
ping 127.0.0.1 -n 3 > nul
rmdir /s /q "$installDir"
del /q "%~f0"
"@ | Set-Content -Path $cleanupScriptPath -Encoding ASCII

        Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `"$cleanupScriptPath`"" -WindowStyle Hidden | Out-Null
    }
}

if ($RemoveUserData -and (Test-Path $userDataDir)) {
    Remove-Item $userDataDir -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output "Uninstalled $appName from $installDir"
if ($RemoveUserData) {
    Write-Output "Removed user data from $userDataDir"
}
else {
    Write-Output "User data was kept at $userDataDir"
}
