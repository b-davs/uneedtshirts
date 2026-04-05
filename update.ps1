$ErrorActionPreference = "Stop"
$repo = "b-davs/uneedtshirts"
$appDir = $PSScriptRoot

Write-Host ""
Write-Host "=== New Order Launcher Updater ===" -ForegroundColor Cyan
Write-Host ""

# Read current version
$versionFile = Join-Path $appDir "version.txt"
if (Test-Path $versionFile) {
  $currentVersion = (Get-Content $versionFile -Raw).Trim()
} else {
  $currentVersion = "0.0.0"
}
Write-Host "Current version: v$currentVersion"

# Check latest release on GitHub
Write-Host "Checking for updates..."
try {
  $release = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" -Headers @{ "User-Agent" = "NewOrderLauncher-Updater" }
} catch {
  Write-Host "Could not reach GitHub. Check your internet connection." -ForegroundColor Red
  Write-Host ""
  Read-Host "Press Enter to close"
  exit 1
}

$latestVersion = $release.tag_name -replace "^v", ""
Write-Host "Latest version:  v$latestVersion"
Write-Host ""

if ($currentVersion -eq $latestVersion) {
  Write-Host "Already up to date!" -ForegroundColor Green
  Write-Host ""
  Read-Host "Press Enter to close"
  exit 0
}

# Find the zip asset
$asset = $release.assets | Where-Object { $_.name -eq "NewOrderLauncher.zip" }
if (-not $asset) {
  Write-Host "No download found for this release." -ForegroundColor Red
  Write-Host ""
  Read-Host "Press Enter to close"
  exit 1
}

Write-Host "Downloading v$latestVersion..." -ForegroundColor Yellow
$tempZip = Join-Path $env:TEMP "NewOrderLauncher_update.zip"
$tempDir = Join-Path $env:TEMP "NewOrderLauncher_update"

# Clean up any previous failed update
if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }

Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $tempZip

Write-Host "Extracting..."
Expand-Archive -Path $tempZip -DestinationPath $tempDir -Force

# Copy files, skipping config.json
$protectedFiles = @("config.json")
$updateFiles = Get-ChildItem -Path $tempDir -File

foreach ($file in $updateFiles) {
  if ($protectedFiles -contains $file.Name) {
    Write-Host "  Skipped $($file.Name) (protected)" -ForegroundColor DarkGray
    continue
  }
  Copy-Item -Path $file.FullName -Destination $appDir -Force
  Write-Host "  Updated $($file.Name)"
}

# Write the new version
Set-Content -Path $versionFile -Value $latestVersion -NoNewline

# Clean up temp files
Remove-Item $tempZip -Force
Remove-Item $tempDir -Recurse -Force

Write-Host ""
Write-Host "Updated to v$latestVersion!" -ForegroundColor Green
Write-Host ""
Read-Host "Press Enter to close"
