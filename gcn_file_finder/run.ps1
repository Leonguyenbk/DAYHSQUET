param(
    [switch]$SkipInstall,
    [switch]$InstallOnly
)

$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvDir = Join-Path $ProjectDir ".venv"
$PythonExe = Join-Path $VenvDir "Scripts\python.exe"
$Requirements = Join-Path $ProjectDir "requirements.txt"
$Stamp = Join-Path $VenvDir ".requirements.sha256"

function Get-Python311 {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $candidate = & py -3.11 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0) { return $candidate }
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $versionOk = & python -c "import sys; print(int(sys.version_info >= (3,11)))"
        if ($versionOk -eq "1") { return (Get-Command python).Source }
    }
    throw "Python 3.11 or newer (64-bit) is required."
}

Set-Location $ProjectDir
if (-not (Test-Path -LiteralPath $PythonExe)) {
    $BasePython = Get-Python311
    Write-Host "Creating virtual environment: $VenvDir"
    & $BasePython -m venv $VenvDir
}

if (-not $SkipInstall) {
    $CurrentHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Requirements).Hash
    $SavedHash = if (Test-Path -LiteralPath $Stamp) { Get-Content -LiteralPath $Stamp -Raw } else { "" }
    if ($CurrentHash -ne $SavedHash.Trim()) {
        & $PythonExe -m pip install --upgrade pip
        & $PythonExe -m pip install -r $Requirements
        if ($LASTEXITCODE -ne 0) { throw "Dependency installation failed." }
        Set-Content -LiteralPath $Stamp -Value $CurrentHash -Encoding ASCII
    }
}

if ($InstallOnly) {
    Write-Host "Environment installed. Run .\download_models.ps1 next."
    exit 0
}

& $PythonExe (Join-Path $ProjectDir "app.py")
if ($LASTEXITCODE -ne 0) { throw "Application exited with code $LASTEXITCODE." }

