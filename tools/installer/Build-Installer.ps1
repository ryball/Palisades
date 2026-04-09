[CmdletBinding()]
param(
    [switch]$SkipPublish,
    [switch]$SkipZip
)

$ErrorActionPreference = 'Stop'

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$publishDir = Join-Path $repoRoot 'artifacts\publish'
$installerDir = Join-Path $repoRoot 'artifacts\installer'
$payloadDir = Join-Path $installerDir 'payload'
$zipPath = Join-Path $repoRoot 'artifacts\DeskHaven-Installer.zip'
$appProject = Join-Path $repoRoot 'Palisades.Application\Palisades.Application.csproj'

Get-Process DeskHaven, Palisades -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

if (-not $SkipPublish) {
    if (Test-Path $publishDir) {
        Remove-Item $publishDir -Recurse -Force
    }

    dotnet publish $appProject -c Release -o $publishDir
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet publish failed with exit code $LASTEXITCODE"
    }
}

if (Test-Path $installerDir) {
    Remove-Item $installerDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $installerDir | Out-Null
if (Test-Path $payloadDir) {
    Remove-Item $payloadDir -Recurse -Force
}
Copy-Item $publishDir $payloadDir -Recurse -Force

Copy-Item (Join-Path $PSScriptRoot 'Install-Palisades.ps1') (Join-Path $installerDir 'Install-DeskHaven.ps1') -Force
Copy-Item (Join-Path $PSScriptRoot 'Uninstall-Palisades.ps1') (Join-Path $installerDir 'Uninstall-DeskHaven.ps1') -Force

@"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-DeskHaven.ps1" %*
"@ | Set-Content -Path (Join-Path $installerDir 'Install-DeskHaven.cmd') -Encoding ASCII

@"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-DeskHaven.ps1" %*
"@ | Set-Content -Path (Join-Path $installerDir 'Uninstall-DeskHaven.cmd') -Encoding ASCII

if (-not $SkipZip) {
    if (Test-Path $zipPath) {
        Remove-Item $zipPath -Force
    }

    Compress-Archive -Path (Join-Path $installerDir '*') -DestinationPath $zipPath -Force
}

Write-Output "Installer prepared at: $installerDir"
if (Test-Path $zipPath) {
    Write-Output "Zip package created at: $zipPath"
}
