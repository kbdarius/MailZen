<#
.SYNOPSIS
    Builds MailZen as a self-contained app and optionally creates an installer.

.DESCRIPTION
    1. Publishes the app as self-contained (no .NET install required for the user)
    2. If Inno Setup is installed, compiles the installer → installer\MailZenSetup.exe
    3. If not, outputs a ready-to-zip publish folder

.EXAMPLE
    .\build-installer.ps1
    .\build-installer.ps1 -SkipInstaller
#>
param(
    [switch]$SkipInstaller
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
if (-not $root) { $root = Split-Path -Parent (Get-Location) }

# If run from the repo root directly
if (Test-Path "$PSScriptRoot\src") { $root = $PSScriptRoot }
if (Test-Path "$root\src\EmailManage.App\EmailManage.App.csproj") {
    # good
} else {
    # Try assuming script is in installer/ folder
    $root = Split-Path -Parent $PSScriptRoot
}

$project  = Join-Path $root "src\EmailManage.App\EmailManage.App.csproj"
$pubDir   = Join-Path $root "publish"
$issFile  = Join-Path $root "installer\MailZen.iss"

Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  MailZen Build & Package Script" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# ── Step 1: Find dotnet ──
$dotnet = $null
foreach ($candidate in @(
    (Join-Path $env:LOCALAPPDATA "dotnet\dotnet.exe"),
    "dotnet"
)) {
    if (Get-Command $candidate -ErrorAction SilentlyContinue) {
        $dotnet = $candidate
        break
    }
}

if (-not $dotnet) {
    Write-Host "ERROR: .NET SDK not found. Install from https://dot.net/download" -ForegroundColor Red
    exit 1
}

Write-Host "[1/3] Publishing self-contained app..." -ForegroundColor Yellow
Write-Host "       dotnet: $dotnet"
Write-Host "       Output: $pubDir"
Write-Host ""

# Clean previous publish
if (Test-Path $pubDir) { Remove-Item $pubDir -Recurse -Force }

& $dotnet publish $project `
    -c Release `
    -r win-x64 `
    --self-contained true `
    -p:PublishReadyToRun=true `
    -p:PublishSingleFile=false `
    -o $pubDir

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: dotnet publish failed with exit code $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}

$exePath = Join-Path $pubDir "MailZen.exe"
if (-not (Test-Path $exePath)) {
    Write-Host "ERROR: MailZen.exe not found in publish output" -ForegroundColor Red
    Write-Host "       Files in publish dir:"
    Get-ChildItem $pubDir -Name | Select-Object -First 15
    exit 1
}

$exeSize = (Get-Item $exePath).Length / 1MB
$totalSize = (Get-ChildItem $pubDir -Recurse | Measure-Object -Property Length -Sum).Sum / 1MB
Write-Host ""
Write-Host "[OK] Published successfully!" -ForegroundColor Green
Write-Host "     MailZen.exe: $([math]::Round($exeSize, 1)) MB"
Write-Host "     Total folder: $([math]::Round($totalSize, 1)) MB"
Write-Host ""

# ── Step 2: Create installer (if Inno Setup available) ──
if ($SkipInstaller) {
    Write-Host "[SKIP] Installer creation skipped (use -SkipInstaller to skip)" -ForegroundColor DarkGray
} else {
    Write-Host "[2/3] Looking for Inno Setup..." -ForegroundColor Yellow

    $iscc = $null
    foreach ($candidate in @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "${env:ProgramFiles(x86)}\Inno Setup 5\ISCC.exe"
    )) {
        if (Test-Path $candidate) {
            $iscc = $candidate
            break
        }
    }

    if ($iscc) {
        Write-Host "       Found: $iscc" -ForegroundColor Gray

        # Check if .iss references an icon file; skip SetupIconFile if icon doesn't exist
        $iconPath = Join-Path $root "src\EmailManage.App\Resources\mailzen.ico"
        if (-not (Test-Path $iconPath)) {
            Write-Host "       NOTE: No icon file found at Resources\mailzen.ico — installer will use default icon" -ForegroundColor DarkYellow
            # Create a temporary .iss without the SetupIconFile line
            $issContent = Get-Content $issFile -Raw
            $tempIss = Join-Path $root "installer\MailZen_temp.iss"
            $issContent -replace 'SetupIconFile=.*\r?\n', '' | Set-Content $tempIss -Encoding UTF8
            $issFile = $tempIss
        }

        & $iscc $issFile
        if ($LASTEXITCODE -eq 0) {
            $installerPath = Join-Path $root "installer\MailZenSetup.exe"
            if (Test-Path $installerPath) {
                $instSize = (Get-Item $installerPath).Length / 1MB
                Write-Host ""
                Write-Host "[OK] Installer created!" -ForegroundColor Green
                Write-Host "     $installerPath ($([math]::Round($instSize, 1)) MB)"
            }
        } else {
            Write-Host "WARNING: Inno Setup compilation failed. You can still distribute the publish folder." -ForegroundColor Yellow
        }

        # Clean up temp file
        if (Test-Path (Join-Path $root "installer\MailZen_temp.iss")) {
            Remove-Item (Join-Path $root "installer\MailZen_temp.iss") -Force
        }
    } else {
        Write-Host "       Inno Setup not found — skipping installer creation" -ForegroundColor DarkYellow
        Write-Host "       Install Inno Setup from: https://jrsoftware.org/isinfo.php" -ForegroundColor DarkGray
        Write-Host "       Or just zip the publish folder to distribute" -ForegroundColor DarkGray
    }
}

# ── Step 3: Summary ──
Write-Host ""
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "  Distribution Options" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Option A: Send the installer"
Write-Host "    → installer\MailZenSetup.exe (if created)"
Write-Host ""
Write-Host "  Option B: Zip the publish folder"
Write-Host "    → Zip the 'publish\' folder and send it"
Write-Host "    → User extracts and runs MailZen.exe"
Write-Host ""
Write-Host "  What the user needs:"
Write-Host "    1. Windows 10/11 (64-bit)"
Write-Host "    2. Microsoft Outlook Desktop (configured with email)"
Write-Host "    3. Ollama (MailZen will offer to install it on first run)"
Write-Host ""
Write-Host "  No .NET installation required (self-contained)!" -ForegroundColor Green
Write-Host ""
