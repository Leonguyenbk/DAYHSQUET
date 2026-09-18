$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PythonExe = Join-Path $ProjectDir ".venv\Scripts\python.exe"

if (-not (Test-Path -LiteralPath $PythonExe)) {
    throw "Missing .venv. Run .\run.ps1 -InstallOnly first."
}

Set-Location $ProjectDir
& $PythonExe (Join-Path $ProjectDir "download_models.py")
if ($LASTEXITCODE -ne 0) { throw "Model download failed. Check network and disk space." }

