# User guide — Print Spooler Queue Cleanup v5

## Main window

```
┌──────────────────────────────────────────────────────────────────────┐
│ 🖨  Print Spooler Queue Cleanup            [● Spooler: Running] [User]│
│    Stuck print jobs, cleared safely.                       [chip row]│
│ [⟳ Refresh] [🗑 Delete selected] [✕ Clear all jobs]  [⏸ Pause] [⟳ Restart spooler]   [⚙] [] [] │
│ ──────────────────────────────────────────────────────────────────┐ │
│ │ ▸ HP LaserJet (2)                                                  │ │
│ │   Document        Owner    Size     Status     Submitted         │ │
│ │   Report.pdf      ANNA     1.20 MB  Printing   16.10.2026 10:31  │ │
│ │   Scan_0034       IVAN     340 KB   Queued     16.10.2026 10:33  │ │
│ └──────────────────────────────────────────────────────────────────┘ │
│ Language [System ▾]  Theme [System ▾]  [⇄] Show tray icon  [⇄] Start with Windows │
│ Jobs: 2 · Printers: 1                            Updated 16.10 10:35 │
└──────────────────────────────────────────────────────────────────────┘
```

* **Refresh (F5)** — re-enumerate printers and jobs immediately. Auto-refresh runs every
  *N* seconds (Settings ▸ Behavior, default 3, 0 = off).
* **Delete selected (Del)** — remove only selected jobs via `JOB_CONTROL_DELETE`.
  Multi-select with Ctrl/Shift. If a job belongs to another user and you are not
  elevated, the app offers elevation (when allowed).
* **Clear all jobs** — confirmation (can be disabled), then: optional spool-file backup →
  API purge per printer → if anything survives: stop spooler, delete only `*.SPL`/`*.SHD`,
  start spooler. Result arrives as a balloon + status line.
* **Pause / Resume queue** — pauses/resumes all printer queues; the button glyph and label
  flip to reflect the current state.
* **Restart spooler** — stops and starts the Print Spooler service (admin; elevation
  offered when allowed).
* **⚙ Settings**, **🌙/☀ theme cycle**, **ℹ About** — see below.
* Header chips: service state (green/amber/red dot), *Administrator* / *Standard user*,
  *Portable mode* when applicable.
* Empty queue → friendly centered empty state instead of a bare list.

### Hotkeys

| Key | Action |
|---|---|
| `F5` | Refresh |
| `Del` | Delete selected jobs |
| `Ctrl+A` | Select all rows |
| `Ctrl+,` | Settings |
| `F1` | Open the wiki knowledge base |
| `Esc` | Close (to tray when enabled) |

## Settings dialog

* **Appearance** — theme (System/Light/Dark), language (System + 30 languages, native names).
* **Behavior** — tray icon; close-button minimizes to tray; start minimized; notifications;
  confirm before clear-all; auto update check; offer elevation; backup before purge;
  start with Windows; auto-refresh period (seconds).
* **Logging** — Off / Errors only / Everything (`%LOCALAPPDATA%\SpoolerCleanup\logs`).
* **Data** — open logs folder; open backups folder; **Restore from backup…** (pick a
  timestamped backup → spooler stop → files copied back → spooler start); **Uninstall…**.

Every switch and combo applies immediately and is persisted to `settings.json`.

## Tray

* **Double-click** — open the main window (and refresh).
* **Right-click** menu: *Open application* (bold default), *Print queue jobs: N* (live,
  greyed info row), *Open print queue* (default printer, or first printer with jobs),
  *Open printer management*, *Clear all jobs*, *Pause/Resume spooler*, *Restart spooler
  service*, *Theme ▸ (System/Light/Dark, checked)*, *Start with Windows* (checked),
  *Check for updates*, *About*, *Exit*.
* Balloons: purge results, restore results, service restart, update available (respect the
  *Show notifications* setting).

## Themes

* *System* reads `HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\AppsUseLightTheme`
  and follows changes live (switching Windows theme flips the app instantly).
* Dark mode uses immersive dark title bar + dark common controls + custom-drawn Fluent
  buttons/switches/chips with the dark palette; light mode mirrors it.
* Per-monitor DPI: move the window between 100 % and 150 % displays — everything rescales.

## Updates

* *About ▸ Check for updates* or tray ▸ *Check for updates* — queries
  `api.github.com/repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest`.
* Silent daily check when *Check for updates automatically* is on.
* Update flow: download to `%TEMP%` → verification → detached child waits for the running
  instance to exit → atomic replace → relaunch. If the file is locked (rare), the replace
  is scheduled for reboot and you are told so.

## Elevation philosophy

The app runs **asInvoker**. Admin rights are needed only for: deleting other users' jobs,
purging across users, restarting the service, all-users install/autostart. When such an
operation hits `ERROR_ACCESS_DENIED`, you get a shield prompt: *Retry as admin* relaunches
just that operation elevated (`/purge` child), the main app stays unprivileged.
