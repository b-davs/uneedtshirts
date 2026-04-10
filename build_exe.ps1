$ErrorActionPreference = "Stop"

python -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name "NewOrderLauncher" `
  --icon "applogo.ico" `
  --add-data "config.example.json;." `
  --hidden-import "win32com" `
  --hidden-import "win32com.client" `
  --hidden-import "tkcalendar" `
  main.py

python -m PyInstaller `
  --noconfirm `
  --onefile `
  --windowed `
  --name "BizactivityWatcher" `
  --icon "applogo.ico" `
  --add-data "config.example.json;." `
  --hidden-import "win32com" `
  --hidden-import "win32com.client" `
  --hidden-import "watchdog" `
  watcher.py

$dist = Join-Path $PSScriptRoot "dist"

if (Test-Path (Join-Path $PSScriptRoot "config.json")) {
  Copy-Item (Join-Path $PSScriptRoot "config.json") $dist -Force
  Write-Host "Copied config.json to dist/"
} else {
  Copy-Item (Join-Path $PSScriptRoot "config.example.json") (Join-Path $dist "config.json") -Force
  Write-Host "Copied config.example.json as config.json to dist/ (edit paths before running)"
}

if (Test-Path (Join-Path $PSScriptRoot "clients.csv")) {
  Copy-Item (Join-Path $PSScriptRoot "clients.csv") $dist -Force
  Write-Host "Copied clients.csv to dist/"
}

Write-Host "Build complete: dist/NewOrderLauncher.exe"
