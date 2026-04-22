<#
.SYNOPSIS
    Downloads all prerequisites for Convert-to-CloudOnly-GUI.ps1 to a
    local folder that can later be copied to an air-gapped network.

.DESCRIPTION
    Run this script on a machine WITH internet access.
    The result (the Offline-Packages folder) is then copied to the
    target machine, where Install-Prerequisites.ps1 is run.

.NOTES
    Run as administrator so the NuGet provider can be saved to
    Program Files (system-wide). Otherwise it is saved per-user.
#>

#Requires -Version 5.1
Set-StrictMode -Off

$ErrorActionPreference = 'Stop'
$OutputDir = Join-Path $PSScriptRoot "Offline-Packages"

function Write-Step ([string]$msg, [string]$color = 'Cyan') {
    Write-Host ""
    Write-Host ">>> $msg" -ForegroundColor $color
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Downloading prerequisites (with internet)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# -- Create output folder --
if (-not (Test-Path $OutputDir)) {
    New-Item $OutputDir -ItemType Directory -Force | Out-Null
}
Write-Host "Packages will be saved to: $OutputDir"

# -- 1. NuGet provider --
Write-Step "Step 1/4: Ensuring the NuGet provider is installed..."
try {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
    Write-Host "    NuGet provider installed/updated." -ForegroundColor Green
} catch {
    Write-Host "    Warning: $_" -ForegroundColor Yellow
}

# Locate and copy the NuGet provider DLL
$nugetDll = Get-ChildItem "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget" `
                -Recurse -Filter "*.dll" -ErrorAction SilentlyContinue |
            Sort-Object FullName | Select-Object -Last 1

if (-not $nugetDll) {
    $nugetDll = Get-ChildItem "$env:LOCALAPPDATA\PackageManagement\ProviderAssemblies\nuget" `
                    -Recurse -Filter "*.dll" -ErrorAction SilentlyContinue |
                Sort-Object FullName | Select-Object -Last 1
}

if ($nugetDll) {
    $nugetDestDir = Join-Path $OutputDir "NuGetProvider\$($nugetDll.Directory.Name)"
    New-Item $nugetDestDir -ItemType Directory -Force | Out-Null
    Copy-Item $nugetDll.FullName $nugetDestDir -Force
    Write-Host "    NuGet provider copied: $($nugetDll.Directory.Name)" -ForegroundColor Green
} else {
    Write-Host "    Could not locate NuGet provider DLL. Continuing without it." -ForegroundColor Yellow
}

# -- 2. Update PowerShellGet --
Write-Step "Step 2/4: Updating PowerShellGet..."
try {
    Install-Module PowerShellGet -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
    Write-Host "    PowerShellGet updated." -ForegroundColor Green
} catch {
    Write-Host "    Warning: $_" -ForegroundColor Yellow
}

# -- 3. Download Microsoft.Graph modules with dependencies --
Write-Step "Step 3/4: Downloading Microsoft.Graph modules (may take a few minutes)..."

$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$modulesDir = Join-Path $OutputDir "Modules"
New-Item $modulesDir -ItemType Directory -Force | Out-Null

foreach ($mod in $graphModules) {
    Write-Host "    Downloading $mod..." -ForegroundColor Gray
    try {
        Save-Module -Name $mod -Path $modulesDir -Force -ErrorAction Stop
        Write-Host "    $mod OK" -ForegroundColor Green
    } catch {
        Write-Host "    FAIL for $mod`: $_" -ForegroundColor Red
    }
}

# -- 4. RSAT information --
Write-Step "Step 4/4: ActiveDirectory module (RSAT)..."
Write-Host @"
    The ActiveDirectory module is a Windows feature and cannot be
    downloaded as a PowerShell package.

    On the target machine (Windows Server):
      Install-WindowsFeature RSAT-AD-PowerShell

    On the target machine (Windows 10/11 client):
      Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

    On an air-gapped network: enable via Windows Features or deploy
    via DISM/SCCM/Intune using an offline source.
"@ -ForegroundColor Yellow

# -- Write manifest --
$manifest = [ordered]@{
    Skapad       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    Dator        = $env:COMPUTERNAME
    PSVersion    = $PSVersionTable.PSVersion.ToString()
    Moduler      = $graphModules
    NuGetVersion = if ($nugetDll) { $nugetDll.Directory.Name } else { "unknown" }
}
$manifest | ConvertTo-Json | Set-Content (Join-Path $OutputDir "manifest.json") -Encoding UTF8

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "  Done!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Copy the entire folder to the target machine:" -ForegroundColor White
Write-Host "  $OutputDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "Then run on the target machine (as administrator):" -ForegroundColor White
Write-Host "  .\Install-Prerequisites.ps1" -ForegroundColor Yellow
Write-Host ""
