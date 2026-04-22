<#
.SYNOPSIS
    Installerar förutsättningar för Convert-to-CloudOnly-GUI.ps1.
    Fungerar både online (PSGallery) och offline (Offline-Packages).

.DESCRIPTION
    Kör detta skript på måldatorn som administratör.

    ONLINE-läge (standard om Offline-Packages saknas):
      Installerar Microsoft.Graph-moduler direkt från PSGallery och
      NuGet-provider via Install-PackageProvider. Kräver internet.

    OFFLINE-läge (aktiveras automatiskt om Offline-Packages\ finns):
      Kräver att Download-Prerequisites.ps1 körts på en internetansluten
      dator och att mappen Offline-Packages\ kopierats hit.

    Kör som administratör.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator
Set-StrictMode -Off

$ErrorActionPreference = 'Stop'
$OfflineDir  = Join-Path $PSScriptRoot "Offline-Packages"
$ModulesDir  = Join-Path $OfflineDir "Modules"
$NuGetDir    = Join-Path $OfflineDir "NuGetProvider"
$InstallScope = "AllUsers"   # AllUsers kräver admin; byt till CurrentUser om det behovs

function Write-Step ([string]$msg, [string]$color = 'Cyan') {
    Write-Host ""
    Write-Host ">>> $msg" -ForegroundColor $color
}
function Write-OK   ([string]$msg) { Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warn ([string]$msg) { Write-Host "    [!!] $msg" -ForegroundColor Yellow }
function Write-Fail ([string]$msg) { Write-Host "    [ERR] $msg" -ForegroundColor Red }

# ── Identifiera läge ──
$Offline = Test-Path $OfflineDir
$ModeLabel = if ($Offline) { 'offline' } else { 'online (PSGallery)' }

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Installation av förutsättningar ($ModeLabel)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

if ($Offline) {
    Write-Host "  Källa: $OfflineDir"
    if (Test-Path (Join-Path $OfflineDir "manifest.json")) {
        $manifest = Get-Content (Join-Path $OfflineDir "manifest.json") -Raw | ConvertFrom-Json
        Write-Host "  Paket skapade: $($manifest.Skapad) på $($manifest.Dator)" -ForegroundColor Gray
    }
} else {
    Write-Host "  Offline-Packages saknas — installerar från PSGallery."
}

# ── 1. NuGet-provider ──
Write-Step "Steg 1/4: Installerar NuGet-provider..."

$systemProviderDir = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget"
$userProviderDir   = "$env:LOCALAPPDATA\PackageManagement\ProviderAssemblies\nuget"

$alreadyInstalled = (Test-Path $systemProviderDir) -or (Test-Path $userProviderDir)

if ($alreadyInstalled) {
    Write-OK "NuGet-provider är redan installerad."
} elseif ($Offline -and (Test-Path $NuGetDir)) {
    try {
        $versionDir = Get-ChildItem $NuGetDir -Directory | Select-Object -First 1
        $dllSource  = Get-ChildItem $versionDir.FullName -Filter "*.dll" | Select-Object -First 1

        $destDir = Join-Path $systemProviderDir $versionDir.Name
        if (-not (Test-Path $destDir)) { New-Item $destDir -ItemType Directory -Force | Out-Null }

        Copy-Item $dllSource.FullName $destDir -Force
        Write-OK "NuGet-provider installerad ($($versionDir.Name))."
    } catch {
        Write-Fail "Kunde inte installera NuGet-provider: $_"
        Write-Warn "Försök manuellt: kopiera DLL till $systemProviderDir\<version>\"
    }
} else {
    # Online: installera från internet
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
        Write-OK "NuGet-provider installerad via internet."
    } catch {
        Write-Warn "Kunde inte installera NuGet-provider: $_"
        Write-Warn "Om modulinstallationen misslyckas, kör: Install-PackageProvider -Name NuGet -Force"
    }
}

# ── 2 & 3. Installera Microsoft.Graph-moduler ──
$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

if ($Offline) {
    Write-Step "Steg 2/4: Registrerar lokal PSRepository..."

    if (-not (Test-Path $ModulesDir)) {
        Write-Fail "Modulkatalogen saknas: $ModulesDir"
        Write-Warn "Kontrollera att Offline-Packages\Modules finns och innehåller modulerna."
        exit 1
    }

    $repoName = "CloudOnlyOfflineRepo"
    if (Get-PSRepository -Name $repoName -ErrorAction SilentlyContinue) {
        Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue
    }

    $repoRegistered = $false
    try {
        Register-PSRepository -Name $repoName `
            -SourceLocation $ModulesDir `
            -InstallationPolicy Trusted `
            -ErrorAction Stop
        Write-OK "Lokal PSRepository registrerad: $repoName"
        $repoRegistered = $true
    } catch {
        Write-Fail "Kunde inte registrera PSRepository: $_"
        Write-Warn "Fallback: kopierar moduler direkt till modulssökväg..."

        $psModulePath = "$env:ProgramFiles\WindowsPowerShell\Modules"
        $copied = 0
        Get-ChildItem $ModulesDir -Directory | ForEach-Object {
            $dest = Join-Path $psModulePath $_.Name
            if (-not (Test-Path $dest)) {
                Copy-Item $_.FullName $dest -Recurse -Force
                $copied++
            }
        }
        Write-Warn "Kopierade $copied modulkataloger direkt till $psModulePath"
        Write-Warn "Hoppar över Install-Module-steget, verifiera manuellt."
    }

    Write-Step "Steg 3/4: Installerar Microsoft.Graph-moduler (offline)..."
    if ($repoRegistered) {
        foreach ($mod in $graphModules) {
            if (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue) {
                Write-OK "$mod är redan installerad."
                continue
            }
            Write-Host "    Installerar $mod..." -ForegroundColor Gray
            try {
                Install-Module -Name $mod `
                    -Repository $repoName `
                    -Scope $InstallScope `
                    -Force -AllowClobber `
                    -ErrorAction Stop
                Write-OK "$mod installerad."
            } catch {
                Write-Fail "Kunde inte installera $mod`: $_"
            }
        }
        Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue
        Write-OK "Tillfällig PSRepository avregistrerad."
    }
} else {
    Write-Step "Steg 2/4: (Hoppas över — ej aktuellt i online-läge)"
    Write-OK "PSGallery används direkt i nästa steg."

    Write-Step "Steg 3/4: Installerar Microsoft.Graph-moduler (online)..."
    foreach ($mod in $graphModules) {
        if (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue) {
            Write-OK "$mod är redan installerad."
            continue
        }
        Write-Host "    Installerar $mod från PSGallery..." -ForegroundColor Gray
        try {
            Install-Module -Name $mod `
                -Repository PSGallery `
                -Scope $InstallScope `
                -Force -AllowClobber `
                -ErrorAction Stop
            Write-OK "$mod installerad."
        } catch {
            Write-Fail "Kunde inte installera $mod`: $_"
        }
    }
}

# ── 4. ActiveDirectory (RSAT) ──
Write-Step "Steg 4/4: Kontrollerar ActiveDirectory-modulen..."

if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
    Write-OK "ActiveDirectory-modulen är installerad."
} else {
    Write-Warn "ActiveDirectory-modulen saknas."
    Write-Host ""
    Write-Host "    Installera RSAT med ett av följande alternativ:" -ForegroundColor Yellow
    Write-Host ""

    # Detektera om det är Server eller Klient
    $osInfo = Get-CimInstance Win32_OperatingSystem
    if ($osInfo.Caption -match "Server") {
        Write-Host "    Windows Server (körs som admin i PowerShell):" -ForegroundColor White
        Write-Host "      Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature" -ForegroundColor Gray
        Write-Host ""
        Write-Host "    Eller via DISM (luftgapat med Windows Server ISO):" -ForegroundColor White
        Write-Host "      DISM /Online /Enable-Feature /FeatureName:RSATClient-Roles-AD-Powershell" -ForegroundColor Gray

        $install = Read-Host "    Vill du försöka installera nu? (J/N)"
        if ($install -match '^[Jj]') {
            try {
                Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -ErrorAction Stop
                Write-OK "RSAT-AD-PowerShell installerad."
            } catch {
                Write-Fail "Installationen misslyckades: $_"
            }
        }
    } else {
        Write-Host "    Windows 10/11 (körs som admin i PowerShell):" -ForegroundColor White
        Write-Host "      Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor Gray
        Write-Host ""
        Write-Host "    Luftgapat alternativ: aktivera via:" -ForegroundColor White
        Write-Host "      Installera valfria funktioner -> RSAT: Active Directory Domain Services" -ForegroundColor Gray
        Write-Host "      (Kräver Windows installation-media om luftgapat)" -ForegroundColor Gray

        $install = Read-Host "    Vill du försöka installera nu (kräver internet om ingen lokal källa finns)? (J/N)"
        if ($install -match '^[Jj]') {
            try {
                Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                Write-OK "ActiveDirectory-modulen installerad."
            } catch {
                Write-Fail "Installationen misslyckades: $_"
                Write-Warn "Aktivera manuellt via Inställningar > Appar > Valfria funktioner > RSAT"
            }
        }
    }
}

# ── Sammanfattning ──
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Verifiering" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

$checks = @(
    @{ Name='ActiveDirectory';                          Label='ActiveDirectory (RSAT)' },
    @{ Name='Microsoft.Graph.Authentication';           Label='Microsoft.Graph.Authentication' },
    @{ Name='Microsoft.Graph.Users';                    Label='Microsoft.Graph.Users' },
    @{ Name='Microsoft.Graph.Identity.DirectoryManagement'; Label='Microsoft.Graph.Identity.DirectoryManagement' }
)

$allOk = $true
foreach ($c in $checks) {
    if (Get-Module -ListAvailable -Name $c.Name -ErrorAction SilentlyContinue) {
        Write-OK $c.Label
    } else {
        Write-Fail "$($c.Label) - EJ INSTALLERAD"
        $allOk = $false
    }
}

Write-Host ""
if ($allOk) {
    Write-Host "  Alla förutsättningar är uppfyllda. Du kan nu köra Launch-GUI.bat." -ForegroundColor Green
} else {
    Write-Host "  En eller flera moduler saknas - se FEL ovan." -ForegroundColor Red
    Write-Host "  Verktyget kan fortfarande köras med Dry Run för moduler som finns." -ForegroundColor Yellow
}
Write-Host ""
