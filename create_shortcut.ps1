param(
  [string]$ExePath = "$(Join-Path $PSScriptRoot 'dist\NewOrderLauncher.exe')",
  [string]$ShortcutName = "New Order Launcher.lnk",
  [string]$IconPath = "$(Join-Path $PSScriptRoot 'applogo.ico')"
)

$desktop = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktop $ShortcutName

if (-not (Test-Path $ExePath)) {
  throw "Executable not found at $ExePath"
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $ExePath
$shortcut.WorkingDirectory = Split-Path $ExePath -Parent
if (Test-Path $IconPath) {
  $shortcut.IconLocation = "$IconPath,0"
}
$shortcut.Save()

Write-Host "Shortcut created: $shortcutPath"
