<#
.SYNOPSIS
    Laddar ned alla förutsättningar för Convert-to-CloudOnly-GUI.ps1
    till en lokal mapp som sedan kan kopieras till ett luftgapat nätverk.

.DESCRIPTION
    Kör detta skript på en dator MED internetåtkomst.
    Resultatet (mappen Offline-Packages) kopieras sedan till
    måldatorn och Install-Prerequisites.ps1 körs där.

.NOTES
    Kör som administratör för att NuGet-providern ska kunna sparas
    till Program Files (systemomfattande). Annars sparas den per användare.
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
Write-Host "  Nedladdning av förutsättningar (med internet)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

# ── Skapa outputmapp ──
if (-not (Test-Path $OutputDir)) {
    New-Item $OutputDir -ItemType Directory -Force | Out-Null
}
Write-Host "Paketen sparas i: $OutputDir"

# ── 1. NuGet-provider ──
Write-Step "Steg 1/4: Säkerställer att NuGet-providern är installerad..."
try {
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
    Write-Host "    NuGet-providern installerad/uppdaterad." -ForegroundColor Green
} catch {
    Write-Host "    Varning: $_" -ForegroundColor Yellow
}

# Hitta och kopiera NuGet-provider-DLL:en
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
    Write-Host "    NuGet-provider kopierad: $($nugetDll.Directory.Name)" -ForegroundColor Green
} else {
    Write-Host "    Kunde inte hitta NuGet-provider-DLL. Installationen fortsätter utan den." -ForegroundColor Yellow
}

# ── 2. Uppdatera PowerShellGet ──
Write-Step "Steg 2/4: Uppdaterar PowerShellGet..."
try {
    Install-Module PowerShellGet -Force -AllowClobber -Scope AllUsers -ErrorAction Stop
    Write-Host "    PowerShellGet uppdaterat." -ForegroundColor Green
} catch {
    Write-Host "    Varning: $_" -ForegroundColor Yellow
}

# ── 3. Ladda ned Microsoft.Graph-moduler med beroenden ──
Write-Step "Steg 3/4: Laddar ned Microsoft.Graph-moduler (kan ta några minuter)..."

$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

$modulesDir = Join-Path $OutputDir "Modules"
New-Item $modulesDir -ItemType Directory -Force | Out-Null

foreach ($mod in $graphModules) {
    Write-Host "    Laddar ned $mod..." -ForegroundColor Gray
    try {
        Save-Module -Name $mod -Path $modulesDir -Force -ErrorAction Stop
        Write-Host "    $mod OK" -ForegroundColor Green
    } catch {
        Write-Host "    FEL för $mod`: $_" -ForegroundColor Red
    }
}

# ── 4. RSAT-information ──
Write-Step "Steg 4/4: ActiveDirectory-modulen (RSAT)..."
Write-Host @"
    ActiveDirectory-modulen är en Windows-funktion och kan inte laddas ned
    som ett PowerShell-paket.

    På måldatorn (Windows Server):
      Install-WindowsFeature RSAT-AD-PowerShell

    På måldatorn (Windows 10/11 klient):
      Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0

    På luftgapat nätverk: aktivera via Windows Features eller distribuera
    via DISM/SCCM/Intune med offline-källa.
"@ -ForegroundColor Yellow

# ── Skriv manifest ──
$manifest = [ordered]@{
    Skapad       = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    Dator        = $env:COMPUTERNAME
    PSVersion    = $PSVersionTable.PSVersion.ToString()
    Moduler      = $graphModules
    NuGetVersion = if ($nugetDll) { $nugetDll.Directory.Name } else { "okänd" }
}
$manifest | ConvertTo-Json | Set-Content (Join-Path $OutputDir "manifest.json") -Encoding UTF8

Write-Host ""
Write-Host "================================================" -ForegroundColor Green
Write-Host "  Klart!" -ForegroundColor Green
Write-Host "================================================" -ForegroundColor Green
Write-Host ""
Write-Host "Kopiera hela mappen till måldatorn:" -ForegroundColor White
Write-Host "  $OutputDir" -ForegroundColor Yellow
Write-Host ""
Write-Host "Kor sedan pa måldatorn (som administratör):" -ForegroundColor White
Write-Host "  .\Install-Prerequisites.ps1" -ForegroundColor Yellow
Write-Host ""
