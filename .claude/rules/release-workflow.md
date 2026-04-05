---
paths:
  - ".github/**"
  - "install.ps1"
  - "build_exe.ps1"
  - "config.example.json"
---

# Release & Delivery Workflow

Development happens on Mac. The app runs on Dan's Windows PC. There is no way to test the full app on Mac (Excel COM and Windows paths), but the logic layer is testable via `python -m pytest`.

## GitHub repo

Private repo at `b-davs/uneedtshirts`. Code is pushed from Mac; a GitHub Action builds the Windows `.exe`.

## How to release an update

```bash
git add <files> && git commit -m "description"
git push
git tag v<MAJOR>.<MINOR>.<PATCH>
git push origin v<MAJOR>.<MINOR>.<PATCH>
```

The `build-release.yml` Action runs on `windows-latest`, builds via PyInstaller, and attaches `NewOrderLauncher.zip` to a GitHub Release. Build takes ~90 seconds.

## What the release zip contains

- `NewOrderLauncher.exe` — the standalone app
- `config.example.json` — template config (not the live `config.json`)
- `applogo.ico` — app icon
- `install.ps1` — first-run setup script
- `create_shortcut.ps1` — standalone shortcut creator

## How Dan installs an update

1. Download `NewOrderLauncher.zip` from the GitHub release (or receive it via email since the repo is private)
2. Extract to a folder (e.g. `C:\NewOrderLauncher\`)
3. First time only: right-click `install.ps1` → "Run with PowerShell" (creates `config.json` from template + desktop shortcut)
4. On updates: replace files but **keep his existing `config.json`** — `install.ps1` will not overwrite it

## config.json

This file has Windows paths specific to Dan's machine. It is `.gitignore`d — only `config.example.json` is committed. Key paths Dan must set:

- `root_paths.clients_root` — where order folders live (e.g. `D:/A Client Sites & Images`)
- `root_paths.templates_root` — where `.xls` templates are (e.g. `C:/Users/dan/Desktop`)
- `templates[].source_path` — full path to each template file

## Important constraints

- Never commit `config.json` — it contains Windows-specific paths
- `clients.csv` contains real client names/addresses — repo must stay private
- PyInstaller build must target Windows, which is why the Action uses `windows-latest`
- `build_exe.ps1` is the legacy local build script; the GitHub Action replaces it but it's kept for manual builds on Dan's PC if needed
