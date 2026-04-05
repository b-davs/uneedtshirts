---
paths:
  - ".github/**"
  - "install.ps1"
  - "update.ps1"
  - "updater.py"
  - "build_exe.ps1"
  - "config.example.json"
  - "version.txt"
---

# Release & Delivery Workflow

Development happens on Mac. The app runs on Dan's Windows PC. There is no way to test the full app on Mac (Excel COM and Windows paths), but the logic layer is testable via `python -m pytest`.

## GitHub repo

Public repo at `b-davs/uneedtshirts`. Code is pushed from Mac; a GitHub Action builds the Windows `.exe`. `clients.csv` is gitignored (contains real client data).

## How to release an update

```bash
git add <files> && git commit -m "description"
git push
git tag v<MAJOR>.<MINOR>.<PATCH>
git push origin v<MAJOR>.<MINOR>.<PATCH>
```

The `build-release.yml` Action runs on `windows-latest`, builds via PyInstaller, and attaches `NewOrderLauncher.zip` to a GitHub Release. The version tag (minus the `v` prefix) is written to `version.txt` inside the zip. Build takes ~90 seconds.

## What the release zip contains

- `NewOrderLauncher.exe` — the standalone app
- `version.txt` — current version number (written by the Action from the git tag)
- `config.example.json` — template config (not the live `config.json`)
- `applogo.ico` — app icon
- `update.ps1` — self-updater script (checks GitHub, downloads, replaces files)
- `install.ps1` — first-run setup script (config + desktop shortcut)
- `create_shortcut.ps1` — standalone shortcut creator

## How Dan installs (first time)

1. Download `NewOrderLauncher.zip` from `https://github.com/b-davs/uneedtshirts/releases/latest`
2. Extract to a permanent folder (e.g. `C:\NewOrderLauncher\`)
3. Right-click `install.ps1` → "Run with PowerShell" (creates `config.json` from template + desktop shortcut)
4. Edit `config.json` to set his Windows paths

## How Dan updates (subsequent releases)

Updates are automatic. On every launch, `updater.py` checks GitHub in a background thread. If a newer version exists, the app shows a dialog: "Version X.Y.Z is available. Update now?" If Dan clicks Yes, it downloads the zip, replaces all files except `config.json`, and offers to restart.

`update.ps1` is also included as a manual fallback — Dan can right-click it → "Run with PowerShell" if the in-app updater fails for any reason.

In both cases, `config.json` is explicitly protected and never overwritten.

## config.json

This file has Windows paths specific to Dan's machine. It is `.gitignore`d — only `config.example.json` is committed. Key paths Dan must set:

- `root_paths.clients_root` — where order folders live (e.g. `D:/A Client Sites & Images`)
- `root_paths.templates_root` — where `.xls` templates are (e.g. `C:/Users/dan/Desktop`)
- `templates[].source_path` — full path to each template file

## Important constraints

- Never commit `config.json` — it contains Windows-specific paths
- Never commit `clients.csv` — contains real client names/addresses
- PyInstaller build must target Windows, which is why the Action uses `windows-latest`
- `build_exe.ps1` is the legacy local build script; the GitHub Action replaces it but it's kept for manual builds on Dan's PC if needed
