[CmdletBinding()]
param(
    [bool]$LaunchAfterInstall = $true,
    [bool]$CreateDesktopShortcut = $true,
    [bool]$CreateStartupShortcut = $false
)

$ErrorActionPreference = 'Stop'

$appName = 'DeskHaven'
$appVersion = '1.1.0'
$publisher = 'StouderIO'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$payloadDir = Join-Path $scriptRoot 'payload'
$installDir = Join-Path $env:LOCALAPPDATA 'Programs\DeskHaven'
$legacyInstallDir = Join-Path $env:LOCALAPPDATA 'Programs\Palisades'
$startMenuDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\DeskHaven'
$legacyStartMenuDir = Join-Path $env:APPDATA 'Microsoft\Windows\Start Menu\Programs\Palisades'
$desktopShortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'DeskHaven.lnk'
$legacyDesktopShortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Palisades.lnk'
$startupShortcutPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'DeskHaven.lnk'
$legacyStartupShortcutPath = Join-Path ([Environment]::GetFolderPath('Startup')) 'Palisades.lnk'
$exePath = Join-Path $installDir 'DeskHaven.exe'
$legacyExePath = Join-Path $installDir 'Palisades.exe'
$uninstallScriptPath = Join-Path $installDir 'Uninstall-DeskHaven.ps1'
$uninstallCmdPath = Join-Path $installDir 'Uninstall-DeskHaven.cmd'
$startMenuShortcutPath = Join-Path $startMenuDir 'DeskHaven.lnk'
$startMenuUninstallShortcutPath = Join-Path $startMenuDir 'Uninstall DeskHaven.lnk'
$registryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\DeskHaven'
$legacyRegistryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Palisades'
$startupRunKeyPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$startupRunValueNames = @('DeskHaven', 'Palisades')

if (-not (Test-Path $payloadDir)) {
    throw "Installer payload was not found at '$payloadDir'. Run tools/installer/Build-Installer.ps1 first."
}

Get-Process DeskHaven, Palisades -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $exePath) -and (Test-Path $legacyExePath)) {
    $exePath = $legacyExePath
}

if (Test-Path $legacyInstallDir) {
    Remove-Item $legacyInstallDir -Recurse -Force -ErrorAction SilentlyContinue
}

New-Item -ItemType Directory -Force -Path $installDir | Out-Null
$robocopyLog = & robocopy $payloadDir $installDir /MIR /R:1 /W:1 /NFL /NDL /NJH /NJS /NP
if ($LASTEXITCODE -gt 7) {
    throw "robocopy failed with exit code $LASTEXITCODE. Output:`n$($robocopyLog -join [Environment]::NewLine)"
}

Copy-Item (Join-Path $scriptRoot 'Uninstall-Palisades.ps1') $uninstallScriptPath -Force
@"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-DeskHaven.ps1" %*
"@ | Set-Content -Path $uninstallCmdPath -Encoding ASCII

if (Test-Path $startupRunKeyPath) {
    foreach ($startupRunValueName in $startupRunValueNames) {
        Remove-ItemProperty -Path $startupRunKeyPath -Name $startupRunValueName -ErrorAction SilentlyContinue
    }
}

function New-WindowsShortcut {
    param(
        [Parameter(Mandatory = $true)][string]$ShortcutPath,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [string]$Arguments = '',
        [string]$WorkingDirectory = '',
        [string]$IconLocation = ''
    )

    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($ShortcutPath)
    $shortcut.TargetPath = $TargetPath
    if ($Arguments) { $shortcut.Arguments = $Arguments }
    if ($WorkingDirectory) { $shortcut.WorkingDirectory = $WorkingDirectory }
    if ($IconLocation) { $shortcut.IconLocation = $IconLocation }
    $shortcut.Save()
}

if (Test-Path $legacyStartMenuDir) {
    Remove-Item $legacyStartMenuDir -Recurse -Force -ErrorAction SilentlyContinue
}

New-Item -ItemType Directory -Force -Path $startMenuDir | Out-Null
New-WindowsShortcut -ShortcutPath $startMenuShortcutPath -TargetPath $exePath -WorkingDirectory $installDir -IconLocation "$exePath,0"
New-WindowsShortcut -ShortcutPath $startMenuUninstallShortcutPath -TargetPath 'powershell.exe' -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$uninstallScriptPath`"" -WorkingDirectory $installDir -IconLocation "$exePath,0"

if ($CreateDesktopShortcut) {
    New-WindowsShortcut -ShortcutPath $desktopShortcutPath -TargetPath $exePath -WorkingDirectory $installDir -IconLocation "$exePath,0"
}
if (Test-Path $legacyDesktopShortcutPath) {
    Remove-Item $legacyDesktopShortcutPath -Force -ErrorAction SilentlyContinue
}

if ($CreateStartupShortcut) {
    New-WindowsShortcut -ShortcutPath $startupShortcutPath -TargetPath $exePath -WorkingDirectory $installDir -IconLocation "$exePath,0"
}
else {
    Remove-Item $startupShortcutPath -Force -ErrorAction SilentlyContinue
}
if (Test-Path $legacyStartupShortcutPath) {
    Remove-Item $legacyStartupShortcutPath -Force -ErrorAction SilentlyContinue
}

if (Test-Path $legacyRegistryPath) {
    Remove-Item $legacyRegistryPath -Recurse -Force -ErrorAction SilentlyContinue
}

New-Item -Path $registryPath -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'DisplayName' -PropertyType String -Value $appName -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'DisplayVersion' -PropertyType String -Value $appVersion -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'Publisher' -PropertyType String -Value $publisher -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'InstallLocation' -PropertyType String -Value $installDir -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'DisplayIcon' -PropertyType String -Value $exePath -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'UninstallString' -PropertyType String -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$uninstallScriptPath`"" -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'QuietUninstallString' -PropertyType String -Value "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$uninstallScriptPath`"" -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'NoModify' -PropertyType DWord -Value 1 -Force | Out-Null
New-ItemProperty -Path $registryPath -Name 'NoRepair' -PropertyType DWord -Value 1 -Force | Out-Null

if ($LaunchAfterInstall -and (Test-Path $exePath)) {
    Start-Process -FilePath $exePath -WorkingDirectory $installDir | Out-Null
}

Write-Output "Installed or updated $appName $appVersion at $installDir"
