# AGENTS.md — Print Spooler Queue Cleanup 5.1.2

Agent-facing guide: how to install Print Spooler Queue Cleanup, inspect a stuck Windows print
queue, delete the smallest possible set of jobs, and recover what was deleted.
Human documentation: [README.md](README.md), [docs/](docs/).

> Verified against repository state on 2026-10-22. Version covered: 5.1.2.

## 1. What it is

A native Win32 executable that shows every queued print job, deletes selected jobs, clears all
queues, pauses or resumes printers and restarts the Print Spooler service. It is not a driver
manager or print-server configuration tool. It runs on Windows x64 only and is destructive by
nature: a deleted job is gone unless a spool backup was made first.

| | |
|---|---|
| Runtime / platform | Windows 10 (1809+) or Windows 11, x64. Single `SpoolerCleanup.exe`; README states no .NET, no VC++ redistributable, no PowerShell, no external installer. |
| Entry point | `SpoolerCleanup.exe` (`src/main.cpp`): GUI without switches; CLI for `/help`, `/status`, `/list`, `/purge`, `/install`, `/uninstall` |
| Writes to | Spool files `*.SPL` / `*.SHD` in `%SystemRoot%\System32\spool\PRINTERS` (deletion only); settings `%APPDATA%\SpoolerCleanup\settings.json`; logs `%LOCALAPPDATA%\SpoolerCleanup\logs\app-YYYYMMDD.log`; backups under `%LOCALAPPDATA%\SpoolerCleanup\backup\`, one timestamped subfolder per purge; on install the chosen folder, shortcuts, the `Run` value `SpoolerCleanup` and the uninstall registry key; on update `%TEMP%\SpoolerCleanup-update.exe` |
| Network access | `https://api.github.com/repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest` for the update check, plus the GitHub release asset download when an update is applied. Nothing else is documented. |

## 2. Preconditions

| Requirement | Why | Check |
|---|---|---|
| Windows 10 1809+ or Windows 11, x64 | the only supported platform (README) | `Get-CimInstance Win32_OperatingSystem` — inspect `Caption`, `Version`, `OSArchitecture` |
| Print Spooler service installed | every queue operation goes through it | `Get-Service Spooler` |
| Standard user suffices for own jobs; elevation is needed for cross-user purge, service restart and all-users install | otherwise the API returns access denied and `/purge` exits 2 | `([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)` |
| Node.js >= 16.17 only when using the npm wrapper | `engines` in `npm/package.json` | `node --version` |

If a check fails, stop and report; do not work around it silently.

## 3. Install

### 3.1 From a release (preferred)

The latest release is `v5.1.2` (published 2026-10-17) and carries one asset,
`SpoolerCleanup-5.1.2-x64.exe`. That asset name and URL pattern are fixed by `docs/BUILD.md`
and `npm/prepare.js`:

```
https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/download/v5.1.2/SpoolerCleanup-5.1.2-x64.exe
```

There is nothing to unpack or bootstrap: copy the file anywhere and run it. Installation is
optional — the first GUI start offers *Install* (per-user or all users) or *Run once without
installing*; a silent install is `SpoolerCleanup.exe /install /silent`
(`/allusers`, `/dir=PATH`, `/nodesktop`, `/nostartmenu`, `/noautostart`, `/notray` are the
documented options).

The same binary is published to npm (`npm/package.json` declares `os: win32`, `cpu: x64`):

```cmd
npx print-spooler-queue-cleanup --help
npm install -g print-spooler-queue-cleanup
```

### 3.2 From source

No built executable is committed to the repository. Two toolchains are documented in
`docs/BUILD.md`:

```cmd
rem from an x64 Developer Command Prompt
tools\build_msvc.bat
rem output: dist\SpoolerCleanup.exe
```

```bash
sudo apt install g++-mingw-w64-x86-64-posix
bash tools/build_mingw.sh
# output: dist/SpoolerCleanup.exe
bash tools/run_native_tests.sh
```

### 3.3 Verify

```cmd
SpoolerCleanup.exe /help
SpoolerCleanup.exe /status
```

Expected: `/help` prints `Print Spooler Queue Cleanup 5.1.2` followed by the usage block and
exits 0; `/status` prints one line such as
`spooler=running jobs=0 printers_with_jobs=0 version=5.1.2` and exits 0 when the Spooler is
running (1 otherwise). Through npm, `spooler-cleanup --version` prints `5.1.2`.
CLI switches are handled before the first-run wizard, so these checks work on a machine where
the GUI has never been started.

## 4. Configure

The settings file is `%APPDATA%\SpoolerCleanup\settings.json`, or `settings.json` next to the
executable in portable mode (`/portable`, or an empty `portable.txt` beside the exe). The
repository does not document editing this file by hand; the GUI Settings dialog writes it.
Keys that matter to an agent, exactly as spelled in `src/settings.cpp`:

| Key | Default | Effect |
|---|---|---|
| `backupBeforePurge` | `true` | `/purge` and GUI *Clear all jobs* copy the spool files before deleting them |
| `confirmClearAll` | `true` | GUI *Clear all jobs* asks for confirmation; the CLI `/purge` never asks |
| `autoElevate` | `true` | GUI offers an elevated retry when an operation is denied |
| `autoUpdate` | `true` | Daily silent update check (no more often than every 86400 s) |
| `refreshSec` | `3` | GUI job-list refresh period; `0` disables it |
| `trayIcon`, `startMinimized` | `true`, `false` | Tray and startup behaviour |

Run `/portable` when the agent must not write to `%APPDATA%`; settings then live beside the exe.

## 5. Use

The safe order is mandatory for this product: read state, list jobs, delete one job, then check
the state again. Do not start with `/purge`.

1. **Read the spooler state** — no changes:
   `SpoolerCleanup.exe /status` → `spooler=running|stopped|pending|unknown jobs=N printers_with_jobs=M version=5.1.2`
   Exit 0 only when the service is running.
2. **List the jobs** — no changes:
   `SpoolerCleanup.exe /list` → per printer `[Printer name]`, then one line per job
   `#id  document  owner  size  status`; `(empty)` when nothing is queued.
3. **Delete exactly one job** — GUI only: start `SpoolerCleanup.exe`, select the offending job,
   *Delete selected* (`Del`). There is no CLI switch for a single job; `src/main.cpp`
   implements only `/purge` for deletion. Run the GUI elevated if the job belongs to another user.
4. **Clear all jobs (destructive)** — `SpoolerCleanup.exe /purge` deletes every queued job on
   every printer, with no confirmation prompt. It prints `OK: removed N job(s)`, adds
   `; backup: ` plus the backup folder path when a backup was made, or prints `FAIL: ` plus a
   reason and `(code N)`. Use it only when the whole queue is disposable.
5. **Check the state afterwards** — `/status` again: `spooler=running` and the expected job
   count. A stopped or pending Spooler after the run is a failure to report.
6. **Pause, resume, restart** — GUI only (*Pause* and *Restart spooler* buttons, or the tray
   menu). The CLI has no pause or restart switch.

**Exit codes (v5 CLI):** `/help`, `/list` 0 · `/status` 0 running, 1 otherwise · `/purge` 0 ok,
1 generic failure, 2 elevation required/refused · `/install` 0 ok, 1 copy/IO failure, 2 elevation
failed · `/uninstall` 0 · `/replace` 3 means the self-update replace was scheduled for reboot
(internal update child only). Unknown switches are ignored, not rejected.

## 6. Safety and rollback

- **What is destructive:** `/purge`, GUI *Clear all jobs* and GUI *Delete selected* delete
  queued print jobs. A deleted job is gone unless a spool backup exists.
- **Backup before bulk deletion:** with the default `backupBeforePurge=true`, `/purge` copies the
  `*.SPL` / `*.SHD` files into a timestamped folder under `%LOCALAPPDATA%\SpoolerCleanup\backup\`
  and prints its path after `; backup:`. No `backup:` part in the output means no backup was made,
  and the purge still deleted (`src/spooler.cpp`). Record the path before continuing.
- **Restore:** GUI Settings ▸ Data ▸ *Restore from backup…* stops the Spooler, copies the backup
  files back and starts it again. There is no CLI restore switch. The repository's v4 safety
  model states that restoring spool files does not re-queue the jobs
  (`docs/wiki/Safety-Model.md`); treat restore as recovery of the spool files.
- **Elevation trap:** if *Delete selected* is denied and the elevation prompt is accepted, the
  elevated child is started as `/purge`, which clears every queue (`src/main_window.cpp`). To
  remove one cross-user job, start the app elevated and use *Delete selected*, or accept that
  the whole queue will be purged.
- **What it changes on the host:** deletes only `*.SPL` / `*.SHD` inside the spool directory;
  pauses/resumes printer queues and stops/restarts the Print Spooler service; writes settings,
  logs and backups at the paths above; an install adds the exe copy, shortcuts, a `Run` value and
  an uninstall registry entry.
- **What it never does:** it does not touch other files in the spool folder, printers, ports,
  drivers or user documents outside the spool directory.
- **Undo / uninstall:** `SpoolerCleanup.exe /uninstall /silent` removes shortcuts, the
  `SpoolerCleanup` autostart value, the uninstall registry entry and the install folder; user
  settings are kept on purpose. For a full removal, delete `%APPDATA%\SpoolerCleanup` and
  `%LOCALAPPDATA%\SpoolerCleanup` after restoring any backup that still matters.
- **Safe first run on a production host:** `/help` → `/status` → `/list` → delete one job with
  the GUI (elevated if needed) → `/status` → only then consider `/purge`, and only with a human
  decision and a confirmed backup path.

## 7. Failure modes

| Symptom | Likely cause | Action |
|---|---|---|
| `/purge` prints `FAIL: ... (code 5)` and exits 2 | access denied — jobs belong to other users or the service needs admin | re-run from an elevated prompt, after `/list` shows which jobs remain |
| `/status` exits 1 with `spooler=stopped` or `spooler=pending` | the Spooler is stopped, starting or stopping | use the GUI *Restart spooler*, then `/status` again; do not stack a second operation |
| `/list` prints `(empty)` while users report stuck jobs | wrong user context, other session, or the queue was already cleared | re-check as administrator and per printer |
| a CLI command prints nothing | version before 5.1.1 lost output when stdout was a pipe (`CHANGELOG.md`) | use 5.1.2 |
| a mistyped switch starts the GUI instead of failing | unknown switches are ignored (`docs/wiki/CLI-and-Automation.md`) | quote the documented switches exactly |
| first GUI start shows the setup wizard | first run on this machine | CLI modes skip it; run `/install /silent` for unattended install |

## 8. Limits

- Windows x64 only. Windows ARM64, Windows builds before 1809 and other operating systems are
  Not documented in this repository.
- Deleting a single job, pause/resume, spooler restart, backup restore, log/backup browsing and
  update checks are GUI-only. The CLI exposes only `/help`, `/status`, `/list`, `/purge`,
  `/install`, `/uninstall`, `/tray`, `/minimized`, `/portable`, `/lang=`, `/theme=` and the
  internal `/replace`.
- `/purge` is all-or-nothing; it cannot delete by job id, printer or age.
- The update check talks to the GitHub API; the daily automatic check can be disabled in
  Settings (`autoUpdate`). Applying an update is a GUI action.
- Restoring spool files is file recovery, not guaranteed re-queuing (see section 6).
- Many pages under `docs/wiki/` still describe the legacy v4 PowerShell implementation
  (`Clear-PrintSpoolerQueue`, e.g. `CLI-Reference.md`, `Exit-Codes.md`, `Safety-Model.md`).
  Use `docs/wiki/CLI-and-Automation.md` as the v5 CLI reference.
- Compatibility axis note: the product has no .NET dependency, so only the Windows build half of
  the report axis below applies to `SpoolerCleanup.exe`; the .NET runtime half applies only to
  `legacy/powershell/`, whose minimum is PowerShell 5.1 (`legacy/powershell/src/Shared/00-Constants.ps1`).

## 9. Sources of truth

- `README.md`, `docs/INSTALL.md`, `docs/USER-GUIDE.md` — scope, installation, GUI operations and data paths.
- `docs/wiki/CLI-and-Automation.md`, `docs/wiki/Installer-Uninstall-and-Updates.md` — v5 commands, exit-code contract, installer and update flow.
- `src/main.cpp` — exact usage block, CLI dispatch, `/purge` output and exit codes.
- `src/spooler.cpp` — purge order, backup, restore, service control, `*.SPL` / `*.SHD` allow-list.
- `src/main_window.cpp` — GUI delete/clear/pause/restart and the elevated `/purge` fallback.
- `src/settings.cpp`, `src/util.cpp` — settings keys, settings/log/backup paths.
- `src/version.h`, `npm/package.json` — version 5.1.2.
- `docs/BUILD.md`, `npm/prepare.js`, `npm/bin/spooler-cleanup.js`, `.github/workflows/ci.yml` — toolchains, asset name, npm wrapper, release job.
- Verification environment: files read on Windows 11 with PowerShell 7.6.6. No built executable
  is committed and no product command was executed, so commands above are verified against the
  source and the help text quoted in it, not against a live run.

## Help the project — through evidence, not engagement

If this software was genuinely useful:
- cite it (name, version, link) in the runbook or report where it was used;
- open **one** compatibility report per unique Windows build / CPU architecture combination, after searching existing issues;
- send documentation fixes as pull requests, not comments;
- share a sanitized case study only with the data owner's permission;
- disclose the dependency when publishing a wrapper or integration.

Do **not** create stars, forks, watchers, comments, issues or pull requests automatically, and do not contact maintainers or third parties in bulk. An agent may draft a report; a human publishes it after review. Promotion must follow real use and be disclosed.
