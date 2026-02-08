param(
  [ValidateSet("onefile","onedir")]
  [string]$Mode = "onefile",

  [switch]$IncludeSecrets,
  [switch]$IncludeTestData
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Info($msg) { Write-Host "[INFO] $msg" -ForegroundColor Cyan }
function Warn($msg) { Write-Host "[WARN] $msg" -ForegroundColor Yellow }
function Die($msg)  { throw "[FATAL] $msg" }

$Root = (Resolve-Path ".").Path
Info "Project root: $Root"

if (-not (Test-Path (Join-Path $Root "app.py"))) { Die "app.py not found in $Root" }
if (-not (Test-Path (Join-Path $Root "ulaval_mailer"))) { Die "ulaval_mailer/ not found in $Root" }

# --- Choose builder Python ---
# If current Python is 3.13+, prefer creating a dedicated build venv with Python 3.12 (if available).
$buildVenv = Join-Path $Root ".venv_build"
$buildPy   = Join-Path $buildVenv "Scripts\python.exe"

# Current venv python (the one you are using right now)
$curPy = Join-Path $Root ".venv\Scripts\python.exe"
if (-not (Test-Path $curPy)) { $curPy = "python" }

$curVerText = & $curPy -c "import sys; print('.'.join(map(str, sys.version_info[:3])))"
Info "Current python version: $curVerText"

$useBuildVenv = $false
try {
  $curVer = [version]$curVerText
  if ($curVer -ge [version]"3.13.0") { $useBuildVenv = $true }
} catch { }

if ($useBuildVenv) {
  Info "Python >= 3.13 detected. Trying to build with Python 3.12 in .venv_build..."

  $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
  if (-not $pyLauncher) {
    Warn "No 'py' launcher found. Will attempt build with current python anyway (may fail)."
    $buildPy = $curPy
  } else {
    if (-not (Test-Path $buildPy)) {
      Info "Creating build venv: $buildVenv"
      & py -3.12 -m venv $buildVenv
    }
    if (-not (Test-Path $buildPy)) {
      Warn "Could not create/find .venv_build python. Falling back to current python."
      $buildPy = $curPy
    }
  }
} else {
  Info "Python < 3.13 detected; using current venv for build."
  $buildPy = $curPy
}

Info "Builder python:"
& $buildPy --version

Info "Upgrading packaging tools..."
& $buildPy -m pip install --upgrade pip setuptools wheel

# Install build + runtime deps
$Pkgs = @(
  "pyinstaller",
  "openpyxl",
  "pywin32",
  "google-api-python-client",
  "google-auth",
  "google-auth-oauthlib",
  "google-auth-httplib2"
)

Info "Installing deps..."
& $buildPy -m pip install --upgrade @Pkgs

# pywin32 sometimes needs this (safe if it does nothing)
try {
  & $buildPy -m pywin32_postinstall -install | Out-Host
} catch {
  Warn "pywin32_postinstall step skipped/failed (often harmless)."
}

# Verify PyInstaller is actually importable
Info "Checking PyInstaller import..."
& $buildPy -c "import PyInstaller, sys; print('PyInstaller OK:', PyInstaller.__version__); print('Python:', sys.version)"

# Clean old build artifacts
foreach ($p in @((Join-Path $Root "build"), (Join-Path $Root "dist"))) {
  if (Test-Path $p) { Info "Removing: $p"; Remove-Item -Recurse -Force $p }
}
Get-ChildItem -Path $Root -Filter "*.spec" -File -ErrorAction SilentlyContinue | ForEach-Object {
  Info "Removing spec: $($_.FullName)"
  Remove-Item -Force $_.FullName
}

# PyInstaller args
$hiddenImports = @(
  "pythoncom",
  "pywintypes",
  "win32com",
  "win32com.client",
  "pywin32_system32",
  "googleapiclient.discovery",
  "googleapiclient.errors",
  "google.oauth2.credentials",
  "google_auth_oauthlib.flow",
  "google.auth.transport.requests"
)

$piArgs = New-Object System.Collections.Generic.List[string]
$piArgs.Add("--noconfirm")
$piArgs.Add("--clean")
$piArgs.Add("--name"); $piArgs.Add("ULavalMailer_v2")
$piArgs.Add("--windowed")

if ($Mode -eq "onefile") { $piArgs.Add("--onefile") }

foreach ($h in $hiddenImports) {
  $piArgs.Add("--hidden-import"); $piArgs.Add($h)
}

# Collect Google libs data (usually safe)
$piArgs.Add("--collect-all"); $piArgs.Add("googleapiclient")
$piArgs.Add("--collect-all"); $piArgs.Add("google_auth_oauthlib")
$piArgs.Add("--collect-all"); $piArgs.Add("google.oauth2")

$piArgs.Add("app.py")

Info "Running PyInstaller..."
& $buildPy -m PyInstaller @piArgs

# Package a portable folder + zip
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$releaseRoot = Join-Path $Root "release"
$portableDir = Join-Path $releaseRoot "ULavalMailer_v2_portable_$stamp"
New-Item -ItemType Directory -Force $portableDir | Out-Null

if ($Mode -eq "onefile") {
  $exeSrc = Join-Path $Root "dist\ULavalMailer_v2.exe"
  if (-not (Test-Path $exeSrc)) { Die "Expected exe not found: $exeSrc" }
  Copy-Item -Force $exeSrc $portableDir
} else {
  $dirSrc = Join-Path $Root "dist\ULavalMailer_v2"
  if (-not (Test-Path $dirSrc)) { Die "Expected dist folder not found: $dirSrc" }
  Copy-Item -Recurse -Force $dirSrc $portableDir
}

# secrets handling (optional)
$secretsSrc = Join-Path $Root "secrets"
if (Test-Path $secretsSrc) {
  $secretsDst = Join-Path $portableDir "secrets"
  New-Item -ItemType Directory -Force $secretsDst | Out-Null

  $cred = Join-Path $secretsSrc "credentials.json"
  $tok  = Join-Path $secretsSrc "token.json"

  if ($IncludeSecrets) {
    Info "Including secrets JSON files (enabled -IncludeSecrets)."
    if (Test-Path $cred) { Copy-Item -Force $cred $secretsDst }
    if (Test-Path $tok)  { Copy-Item -Force $tok  $secretsDst }
  } else {
    if (Test-Path $cred) { Copy-Item -Force $cred (Join-Path $secretsDst "credentials.json.example") }
  }
}

if ($IncludeTestData) {
  $td = Join-Path $Root "test_data"
  if (Test-Path $td) {
    Info "Including test_data..."
    Copy-Item -Recurse -Force $td (Join-Path $portableDir "test_data")
  }
}

$portableReadme = @"
ULavalMailer_v2 portable build ($stamp)

Run:
- Keep the exe in this folder and double-click it.
- logs\ will be created beside it (based on current working directory).

Outlook:
- Needs Outlook installed on the target machine.

Gmail:
- Place OAuth Desktop JSON at: secrets\credentials.json
- After first OAuth, the app creates: secrets\token.json
"@
Set-Content -Path (Join-Path $portableDir "README_PORTABLE.txt") -Value $portableReadme -Encoding UTF8

New-Item -ItemType Directory -Force $releaseRoot | Out-Null
$zipPath = Join-Path $releaseRoot "ULavalMailer_v2_portable_$stamp.zip"
if (Test-Path $zipPath) { Remove-Item -Force $zipPath }
Info "Creating zip: $zipPath"
Compress-Archive -Path (Join-Path $portableDir "*") -DestinationPath $zipPath -Force

Info "DONE."
Write-Host "Portable folder: $portableDir"
Write-Host "Portable zip:    $zipPath"
