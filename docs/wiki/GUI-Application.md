# GUI application (v5)

The native v5 build is a single `SpoolerCleanup.exe` (Win32/C++, no .NET). This page maps
every surface of the graphical interface.

## Main window

* **Header** — product icon, title, tagline; right-aligned chips: spooler service state
  (green = running, amber = pending, red = stopped), current privilege
  (*Administrator* / *Standard user*), *Portable mode* when active.
* **Toolbar** — `Refresh`, `Delete selected`, `Clear all jobs` (accent-danger),
  `Pause/Resume queue` (glyph and label reflect state), `Restart spooler`; on the right:
  `Settings`, `Theme cycle`, `About`.
* **Job list** — ListView with one group per printer that has jobs; columns: Document,
  Owner, Size, Status, Submitted. Multi-select (Ctrl/Shift). Auto-refresh every N seconds
  (default 3; 0 disables). Empty queue shows a centered empty state.
* **Quick settings row** — Language combo (System + 30 languages), Theme combo
  (System/Light/Dark), *Show tray icon* switch, *Start with Windows* switch.
* **Status bar** — `Jobs: N · Printers: M` on the left, `Updated <time>` on the right.

## Hotkeys

`F5` refresh · `Del` delete selected · `Ctrl+A` select all · `Ctrl+,` settings ·
`F1` knowledge base · `Esc` close (to tray when enabled).

## Settings dialog

Sections: **Appearance** (theme, language), **Behavior** (tray icon, close-to-tray,
start minimized, notifications, confirm clear-all, auto update check, offer elevation,
backup before purge, start with Windows, auto-refresh seconds), **Logging**
(off/errors/everything), **Data** (open logs, open backups, restore from backup,
uninstall). All changes apply instantly and persist to `settings.json`.

## About dialog

Version (quad + 64-bit), author **Mikhail Deynekin <icons@Deynekin.com>**,
© 2026, clickable links: [website](https://Deynekin.com),
[source repository](https://github.com/paulmann/Print-Spooler-Queue-Cleanup),
[license — PolyForm Noncommercial 1.0.0](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE),
[knowledge base (this wiki)](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki);
buttons *Check for updates* and *Copy diagnostics* (version, OS build, elevation, language,
theme, settings path — for issue reports).

## Restore from backup

Settings ▸ Data ▸ *Restore from backup…* lists timestamped backups
(`%LOCALAPPDATA%\SpoolerCleanup\backup\<stamp>`) with file count and size. Restore stops
the spooler, copies `*.SPL`/`*.SHD` back, starts the spooler and refreshes the list.

See also: [Tray icon and notifications](Tray-Icon-and-Notifications),
[Themes and languages](Themes-and-Languages),
[Installer, uninstall and updates](Installer-Uninstall-and-Updates),
[CLI and automation](CLI-and-Automation).
