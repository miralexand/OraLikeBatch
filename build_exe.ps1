$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$venvDir = Join-Path $PSScriptRoot ".venv-build"
$venvPython = Join-Path $venvDir "Scripts\python.exe"

if (-not (Test-Path -LiteralPath $venvPython)) {
    Write-Host "[0/3] Creating isolated build environment..." -ForegroundColor Cyan
    python -m venv $venvDir
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to create virtual environment."
    }
}

Write-Host "[1/3] Installing runtime dependencies and PyInstaller..." -ForegroundColor Cyan
& $venvPython -m pip install --disable-pip-version-check -r requirements.txt pyinstaller
if ($LASTEXITCODE -ne 0) {
    throw "Dependency installation failed. Check network and Python environment."
}

Write-Host "[2/3] Building with PyInstaller..." -ForegroundColor Cyan
& $venvPython -m PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --windowed `
    --name OracleBatchQuery `
    --collect-all oracledb `
    --hidden-import pandas.io.excel._openpyxl `
    --hidden-import openpyxl.cell._writer `
    oracle_batch_query.py
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed."
}

$exePath = Join-Path $PSScriptRoot "dist\OracleBatchQuery.exe"
Write-Host "[3/3] Build finished: $exePath" -ForegroundColor Green
