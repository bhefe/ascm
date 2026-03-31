# Build NSIS Installer Script
# This script prepares the NSIS installer for release

Write-Host "`n🔨 ASCM NSIS Installer Builder`n" -ForegroundColor Cyan

# Check if NSIS is installed
$nsisPath = "C:\Program Files\NSIS"
if (-not (Test-Path $nsisPath)) {
    $nsisPath = "C:\Program Files (x86)\NSIS"
}

if (-not (Test-Path $nsisPath)) {
    Write-Host "❌ NSIS not found!" -ForegroundColor Red
    Write-Host "Please install NSIS from: https://nsis.sourceforge.io/Download" -ForegroundColor Yellow
    Write-Host "`nAfter installation, run this script again.`n" -ForegroundColor Yellow
    exit 1
}

$makensisExe = "$nsisPath\makensis.exe"

if (-not (Test-Path $makensisExe)) {
    Write-Host "❌ makensis.exe not found!" -ForegroundColor Red
    Write-Host "Expected location: $makensisExe" -ForegroundColor Yellow
    exit 1
}

Write-Host "✅ NSIS found at: $nsisPath`n" -ForegroundColor Green

# Verify required files
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$requiredFiles = @(
    "dist\Scan Software.exe",
    "browser_extension",
    "README.md",
    "requirements.txt",
    "installer\ascm.nsi"
)

Write-Host "Checking required files...`n" -ForegroundColor Gray

$allFound = $true
foreach ($file in $requiredFiles) {
    $path = Join-Path $scriptDir $file
    if (Test-Path $path) {
        Write-Host "  ✓ $file" -ForegroundColor Green
    } else {
        Write-Host "  ✗ $file (MISSING!)" -ForegroundColor Red
        $allFound = $false
    }
}

if (-not $allFound) {
    Write-Host "`n❌ Some required files are missing!`n" -ForegroundColor Red
    exit 1
}

Write-Host "`n✅ All required files found!`n" -ForegroundColor Green

# Ask for version
$version = Read-Host "Enter version number (default: 1.5.21)"
if (-not $version) { $version = "1.5.21" }

# Build the installer
$nsiFile = Join-Path $scriptDir "installer\ascm.nsi"
$outputDir = Join-Path $scriptDir "installer\output"

# Create output directory
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

Write-Host "Building NSIS installer...`n" -ForegroundColor Cyan
Write-Host "Using: $makensisExe`n" -ForegroundColor Gray

# Run makensis
Push-Location $scriptDir
& $makensisExe /V4 /DPRODUCT_VERSION=$version $nsiFile
$buildResult = $?
Pop-Location

if ($buildResult) {
    Write-Host "`n✅ Installer built successfully!`n" -ForegroundColor Green
    
    # Find the built installer
    $installerPath = Join-Path $scriptDir "ASCM-Installer-$version.exe"
    
    if (Test-Path $installerPath) {
        $fileSize = (Get-Item $installerPath).Length / 1MB
        Write-Host "Installer Details:" -ForegroundColor Cyan
        Write-Host "  Location: $installerPath" -ForegroundColor White
        Write-Host "  Size: $([Math]::Round($fileSize, 2)) MB" -ForegroundColor White
        Write-Host "  Version: $version" -ForegroundColor White
        Write-Host "`nDistribution ready! Users can run this .exe to install ASCM.`n" -ForegroundColor Green
    }
} else {
    Write-Host "`n❌ Build failed!`n" -ForegroundColor Red
    exit 1
}
