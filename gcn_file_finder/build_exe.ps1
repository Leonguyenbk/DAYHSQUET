$ErrorActionPreference = "Stop"
$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BuildVenv = Join-Path $ProjectDir ".venv-build"
$PythonExe = Join-Path $BuildVenv "Scripts\python.exe"
$ModelsDir = Join-Path $ProjectDir "models\paddleocr"
$DistDir = Join-Path $ProjectDir "dist"
$WorkDir = Join-Path $ProjectDir "build"

function Get-Python311 {
    $RuntimeVenvPython = Join-Path $ProjectDir ".venv\Scripts\python.exe"
    if (Test-Path -LiteralPath $RuntimeVenvPython) {
        $versionOk = & $RuntimeVenvPython -c "import sys; print(int(sys.version_info >= (3,11)))"
        if ($versionOk -eq "1") { return $RuntimeVenvPython }
    }
    if (Get-Command py -ErrorAction SilentlyContinue) {
        $candidate = & py -3.11 -c "import sys; print(sys.executable)" 2>$null
        if ($LASTEXITCODE -eq 0) { return $candidate }
    }
    if (Get-Command python -ErrorAction SilentlyContinue) {
        $versionOk = & python -c "import sys; print(int(sys.version_info >= (3,11)))"
        if ($versionOk -eq "1") { return (Get-Command python).Source }
    }
    throw "Python 3.11 or newer (64-bit) is required for build."
}

Set-Location $ProjectDir
foreach ($name in @("det", "rec", "cls")) {
    $folder = Join-Path $ModelsDir $name
    if (-not (Test-Path -LiteralPath $folder) -or -not (Get-ChildItem -LiteralPath $folder -Force | Select-Object -First 1)) {
        throw "Missing OCR model: $name. Run .\download_models.ps1 first."
    }
}

if (-not (Test-Path -LiteralPath $PythonExe)) {
    $BasePython = Get-Python311
    & $BasePython -m venv $BuildVenv
}

try {
    & $PythonExe -m pip install --upgrade pip
    & $PythonExe -m pip install -r (Join-Path $ProjectDir "requirements.txt")
    if ($LASTEXITCODE -ne 0) { throw "Build dependency installation failed." }

    & $PythonExe -m pytest (Join-Path $ProjectDir "tests") -q
    if ($LASTEXITCODE -ne 0) { throw "Pytest failed; EXE build was cancelled." }

    # onedir is more reliable than onefile for PaddleOCR and keeps offline models beside the executable.
    & $PythonExe -m PyInstaller --noconfirm --clean --onedir --windowed `
        --name "GCNFileFinder" `
        --paths (Split-Path -Parent $ProjectDir) `
        --distpath $DistDir --workpath $WorkDir `
        --collect-all customtkinter --collect-all paddleocr --collect-all paddlex --collect-all paddle `
        --collect-all cv2 --collect-all fitz `
        --copy-metadata imagesize --copy-metadata pyclipper --copy-metadata pypdfium2 `
        --copy-metadata python-bidi --copy-metadata shapely `
        --add-data "$ModelsDir;models/paddleocr" `
        --add-data "$(Join-Path $ProjectDir 'config.example.json');." `
        (Join-Path $ProjectDir "app.py")
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }
    Write-Host "Build completed: $(Join-Path $DistDir 'GCNFileFinder\GCNFileFinder.exe')" -ForegroundColor Green
}
catch {
    Write-Error $_
    throw
}
