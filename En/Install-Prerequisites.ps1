<#
.SYNOPSIS
    Installs prerequisites for Convert-to-CloudOnly-GUI.ps1.
    Works both online (PSGallery) and offline (Offline-Packages).

.DESCRIPTION
    Run this script on the target machine as administrator.

    ONLINE mode (default if Offline-Packages is missing):
      Installs Microsoft.Graph modules directly from PSGallery and
      the NuGet provider via Install-PackageProvider. Requires internet.

    OFFLINE mode (automatically enabled if Offline-Packages\ exists):
      Requires that Download-Prerequisites.ps1 has been run on an
      internet-connected machine and that the Offline-Packages\ folder
      has been copied here.

    Run as administrator.
#>

#Requires -Version 5.1
#Requires -RunAsAdministrator
Set-StrictMode -Off

$ErrorActionPreference = 'Stop'
$OfflineDir  = Join-Path $PSScriptRoot "Offline-Packages"
$ModulesDir  = Join-Path $OfflineDir "Modules"
$NuGetDir    = Join-Path $OfflineDir "NuGetProvider"
$InstallScope = "AllUsers"   # AllUsers requires admin; change to CurrentUser if needed

function Write-Step ([string]$msg, [string]$color = 'Cyan') {
    Write-Host ""
    Write-Host ">>> $msg" -ForegroundColor $color
}
function Write-OK   ([string]$msg) { Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warn ([string]$msg) { Write-Host "    [!!] $msg" -ForegroundColor Yellow }
function Write-Fail ([string]$msg) { Write-Host "    [ERR] $msg" -ForegroundColor Red }

# -- Detect mode --
$Offline = Test-Path $OfflineDir
$ModeLabel = if ($Offline) { 'offline' } else { 'online (PSGallery)' }

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Installing prerequisites ($ModeLabel)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan

if ($Offline) {
    Write-Host "  Source: $OfflineDir"
    if (Test-Path (Join-Path $OfflineDir "manifest.json")) {
        $manifest = Get-Content (Join-Path $OfflineDir "manifest.json") -Raw | ConvertFrom-Json
        Write-Host "  Packages created: $($manifest.Skapad) on $($manifest.Dator)" -ForegroundColor Gray
    }
} else {
    Write-Host "  Offline-Packages missing - installing from PSGallery."
}

# -- 1. NuGet provider --
Write-Step "Step 1/4: Installing NuGet provider..."

$systemProviderDir = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget"
$userProviderDir   = "$env:LOCALAPPDATA\PackageManagement\ProviderAssemblies\nuget"

$alreadyInstalled = (Test-Path $systemProviderDir) -or (Test-Path $userProviderDir)

if ($alreadyInstalled) {
    Write-OK "NuGet provider is already installed."
} elseif ($Offline -and (Test-Path $NuGetDir)) {
    try {
        $versionDir = Get-ChildItem $NuGetDir -Directory | Select-Object -First 1
        $dllSource  = Get-ChildItem $versionDir.FullName -Filter "*.dll" | Select-Object -First 1

        $destDir = Join-Path $systemProviderDir $versionDir.Name
        if (-not (Test-Path $destDir)) { New-Item $destDir -ItemType Directory -Force | Out-Null }

        Copy-Item $dllSource.FullName $destDir -Force
        Write-OK "NuGet provider installed ($($versionDir.Name))."
    } catch {
        Write-Fail "Could not install NuGet provider: $_"
        Write-Warn "Try manually: copy the DLL to $systemProviderDir\<version>\"
    }
} else {
    # Online: install from internet
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop | Out-Null
        Write-OK "NuGet provider installed via internet."
    } catch {
        Write-Warn "Could not install NuGet provider: $_"
        Write-Warn "If module installation fails, run: Install-PackageProvider -Name NuGet -Force"
    }
}

# -- 2 & 3. Install Microsoft.Graph modules --
$graphModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Identity.DirectoryManagement'
)

if ($Offline) {
    Write-Step "Step 2/4: Registering local PSRepository..."

    if (-not (Test-Path $ModulesDir)) {
        Write-Fail "Modules directory missing: $ModulesDir"
        Write-Warn "Verify that Offline-Packages\Modules exists and contains the modules."
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
        Write-OK "Local PSRepository registered: $repoName"
        $repoRegistered = $true
    } catch {
        Write-Fail "Could not register PSRepository: $_"
        Write-Warn "Fallback: copying modules directly to module path..."

        $psModulePath = "$env:ProgramFiles\WindowsPowerShell\Modules"
        $copied = 0
        Get-ChildItem $ModulesDir -Directory | ForEach-Object {
            $dest = Join-Path $psModulePath $_.Name
            if (-not (Test-Path $dest)) {
                Copy-Item $_.FullName $dest -Recurse -Force
                $copied++
            }
        }
        Write-Warn "Copied $copied module directories directly to $psModulePath"
        Write-Warn "Skipping Install-Module step; verify manually."
    }

    Write-Step "Step 3/4: Installing Microsoft.Graph modules (offline)..."
    if ($repoRegistered) {
        foreach ($mod in $graphModules) {
            if (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue) {
                Write-OK "$mod is already installed."
                continue
            }
            Write-Host "    Installing $mod..." -ForegroundColor Gray
            try {
                Install-Module -Name $mod `
                    -Repository $repoName `
                    -Scope $InstallScope `
                    -Force -AllowClobber `
                    -ErrorAction Stop
                Write-OK "$mod installed."
            } catch {
                Write-Fail "Could not install $mod`: $_"
            }
        }
        Unregister-PSRepository -Name $repoName -ErrorAction SilentlyContinue
        Write-OK "Temporary PSRepository unregistered."
    }
} else {
    Write-Step "Step 2/4: (Skipped - not applicable in online mode)"
    Write-OK "PSGallery will be used directly in the next step."

    Write-Step "Step 3/4: Installing Microsoft.Graph modules (online)..."
    foreach ($mod in $graphModules) {
        if (Get-Module -ListAvailable -Name $mod -ErrorAction SilentlyContinue) {
            Write-OK "$mod is already installed."
            continue
        }
        Write-Host "    Installing $mod from PSGallery..." -ForegroundColor Gray
        try {
            Install-Module -Name $mod `
                -Repository PSGallery `
                -Scope $InstallScope `
                -Force -AllowClobber `
                -ErrorAction Stop
            Write-OK "$mod installed."
        } catch {
            Write-Fail "Could not install $mod`: $_"
        }
    }
}

# -- 4. ActiveDirectory (RSAT) --
Write-Step "Step 4/4: Checking the ActiveDirectory module..."

if (Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue) {
    Write-OK "ActiveDirectory module is installed."
} else {
    Write-Warn "ActiveDirectory module is missing."
    Write-Host ""
    Write-Host "    Install RSAT using one of the following options:" -ForegroundColor Yellow
    Write-Host ""

    # Detect Server vs. client
    $osInfo = Get-CimInstance Win32_OperatingSystem
    if ($osInfo.Caption -match "Server") {
        Write-Host "    Windows Server (run as admin in PowerShell):" -ForegroundColor White
        Write-Host "      Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature" -ForegroundColor Gray
        Write-Host ""
        Write-Host "    Or via DISM (air-gapped with Windows Server ISO):" -ForegroundColor White
        Write-Host "      DISM /Online /Enable-Feature /FeatureName:RSATClient-Roles-AD-Powershell" -ForegroundColor Gray

        $install = Read-Host "    Do you want to attempt installation now? (Y/N)"
        if ($install -match '^[Yy]') {
            try {
                Install-WindowsFeature RSAT-AD-PowerShell -IncludeAllSubFeature -ErrorAction Stop
                Write-OK "RSAT-AD-PowerShell installed."
            } catch {
                Write-Fail "Installation failed: $_"
            }
        }
    } else {
        Write-Host "    Windows 10/11 (run as admin in PowerShell):" -ForegroundColor White
        Write-Host "      Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ForegroundColor Gray
        Write-Host ""
        Write-Host "    Air-gapped alternative: enable via:" -ForegroundColor White
        Write-Host "      Optional features -> RSAT: Active Directory Domain Services" -ForegroundColor Gray
        Write-Host "      (Requires Windows installation media if air-gapped)" -ForegroundColor Gray

        $install = Read-Host "    Do you want to attempt installation now (requires internet if no local source exists)? (Y/N)"
        if ($install -match '^[Yy]') {
            try {
                Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop
                Write-OK "ActiveDirectory module installed."
            } catch {
                Write-Fail "Installation failed: $_"
                Write-Warn "Enable manually via Settings > Apps > Optional features > RSAT"
            }
        }
    }
}

# -- Summary --
Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  Verification" -ForegroundColor Cyan
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
        Write-Fail "$($c.Label) - NOT INSTALLED"
        $allOk = $false
    }
}

Write-Host ""
if ($allOk) {
    Write-Host "  All prerequisites satisfied. You can now run Launch-GUI.bat." -ForegroundColor Green
} else {
    Write-Host "  One or more modules are missing - see ERR above." -ForegroundColor Red
    Write-Host "  The tool can still run in Dry Run for modules that are present." -ForegroundColor Yellow
}
Write-Host ""
