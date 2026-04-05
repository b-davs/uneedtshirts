$ErrorActionPreference = "Stop"

# Create config.json from example if it doesn't exist yet
if (-not (Test-Path (Join-Path $PSScriptRoot "config.json"))) {
  Copy-Item (Join-Path $PSScriptRoot "config.example.json") (Join-Path $PSScriptRoot "config.json") -Force
  Write-Host "Created config.json — edit the paths in this file before launching."
} else {
  Write-Host "Existing config.json preserved."
}

# Create desktop shortcut
$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop "New Order Launcher.lnk"
$exePath = Join-Path $PSScriptRoot "NewOrderLauncher.exe"
$iconPath = Join-Path $PSScriptRoot "applogo.ico"

if (-not (Test-Path $exePath)) {
  throw "NewOrderLauncher.exe not found — make sure you extracted the full zip."
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = $PSScriptRoot
if (Test-Path $iconPath) {
  $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Save()

Write-Host ""
Write-Host "Done! Desktop shortcut created."
Write-Host "Make sure config.json has the correct paths before launching."
