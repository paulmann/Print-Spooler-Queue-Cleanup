# Migration guide: PowerShell v4 → native v5

v5 is a **rewrite**, not a patch: the safety model and the mental model survive, the
runtime does not. The v4 PowerShell tree remains fully functional under
[`legacy/powershell/`](../../legacy/powershell/) with its own docs and tests.

## Feature map

| v4 (PowerShell) | v5 (native) | Notes |
|---|---|---|
| `Clear-PrintSpoolerQueue.ps1` interactive/console | GUI main window + `/purge` CLI | same guarded algorithm: backup → API purge → service restart + `*.SPL/*.SHD` sweep |
| `-WhatIf` / safety prompts | confirmation dialog + *Ask before clearing all jobs* setting | plus backup-before-purge default-on |
| Backup/restore switches | Settings ▸ Data ▸ Restore dialog | backups listed with time/size/file count |
| Structured console reports | `/list`, `/status` machine-friendly output | exit codes 0/1/2 |
| Config PSD1 | `settings.json` (%APPDATA% or portable) | schema documented in INSTALL.md |
| Self-update (`-Update`) | tray/About update check + in-place replace | GitHub Releases, `*x64*.exe` asset |
| Event-log / report files | log file (`%LOCALAPPDATA%\…\logs`) | level configurable |
| Fleet/remote management | **not in v5** | use legacy v4 for remote/fleet scenarios |
| PS 5.1 compatibility layer | n/a | v5 needs nothing but Windows 10/11 x64 |

## Automation migration

```powershell
# v4
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup

# v5 equivalent
SpoolerCleanup.exe /purge        # backup governed by settings (backupBeforePurge)
if ($LASTEXITCODE -eq 2) { Start-Process SpoolerCleanup.exe -Verb RunAs -ArgumentList '/purge' -Wait }
```

Deployment: v4 = copy script + set policy; v5 = copy exe (portable) or
`SpoolerCleanup.exe /install /silent` (registered uninstall, autostart optional).

## Why the rewrite

* Zero interpreter dependency and instant start (no PS host spin-up).
* Visual queue inspection — the #1 support ask ("what is actually stuck?").
* Tray guardian with live counter for support engineers.
* Store distribution path (MSIX full-trust) impossible for a loose .ps1.
