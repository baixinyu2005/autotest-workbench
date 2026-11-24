# Build a portable Windows package for AutoTest Workbench.
# Usage: powershell -ExecutionPolicy Bypass -File .\build.ps1

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$python = Join-Path $PSScriptRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $python)) {
    throw "Missing virtual environment. Run: python -m venv .venv"
}

$dist = Join-Path $PSScriptRoot "dist"
$build = Join-Path $PSScriptRoot "build"
$appDir = Join-Path $dist "AutoTestWorkbench"
$zip = Join-Path $dist "AutoTestWorkbench-portable.zip"

if (Test-Path -LiteralPath $appDir) { Remove-Item -LiteralPath $appDir -Recurse -Force }
if (Test-Path -LiteralPath $zip) { Remove-Item -LiteralPath $zip -Force }

$args = @(
    "-m", "PyInstaller",
    "--noconfirm",
    "--clean",
    "--windowed",
    "--onedir",
    "--name", "AutoTestWorkbench",
    "--collect-all", "playwright",
    "--collect-all", "openpyxl",
    "--paths", $PSScriptRoot,
    (Join-Path $PSScriptRoot "app.py")
)

& $python @args
if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE" }

Copy-Item -LiteralPath (Join-Path $PSScriptRoot "README.md") -Destination $appDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "QUICKSTART.txt") -Destination $appDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "ERP_ACCESS_CHECKLIST.txt") -Destination $appDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "demo") -Destination $appDir -Recurse -Force

Compress-Archive -Path $appDir -DestinationPath $zip -CompressionLevel Optimal

$sizeMb = [math]::Round((Get-Item -LiteralPath $zip).Length / 1MB, 1)
Write-Host "Build complete: $zip ($sizeMb MB)" -ForegroundColor Green
Write-Host "Executable: $(Join-Path $appDir 'AutoTestWorkbench.exe')" -ForegroundColor Green
