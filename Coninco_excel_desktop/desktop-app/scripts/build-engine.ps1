$ErrorActionPreference = "Stop"

$Root = (Resolve-Path (Join-Path $PSScriptRoot "../..")).Path
$Venv = Join-Path $Root "python-engine\.venv"
$Output = Join-Path $Root "desktop-app\src-tauri\engine"
$Python = Join-Path $Venv "Scripts\python.exe"
$PyInstaller = Join-Path $Venv "Scripts\pyinstaller.exe"

if (-not (Test-Path $Python)) {
    python -m venv $Venv
}

& $Python -m pip install -r (Join-Path $Root "python-engine\requirements.txt")
New-Item -ItemType Directory -Force $Output | Out-Null
Remove-Item (Join-Path $Output "excel-auditor-engine.exe") -ErrorAction SilentlyContinue

& $PyInstaller `
    --noconfirm `
    --clean `
    --onefile `
    --name excel-auditor-engine `
    --paths (Join-Path $Root "python-engine\src") `
    --collect-all pandas `
    --collect-all openpyxl `
    --collect-all xlrd `
    --distpath $Output `
    --workpath (Join-Path $Root "target\pyinstaller-work") `
    --specpath (Join-Path $Root "target") `
    (Join-Path $Root "python-engine\engine_entry.py")

Write-Host "Built $Output\excel-auditor-engine.exe"
