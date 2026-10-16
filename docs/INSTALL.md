# Installation, first run, portable mode, uninstall

Print Spooler Queue Cleanup v5 ships as **one executable** and installs **itself** — there is
no MSI/EXE bootstrap, no WiX, no NSIS: the installer is a wizard inside the same binary.

## First run on a new machine

Start `SpoolerCleanup.exe` (double-click). Because nothing is installed yet, the setup
wizard opens:

1. **Welcome** — choose *Install* (recommended) or *Run once without installing*.
   *Run once* skips everything, remembers the choice and runs in portable-ish mode
   (settings still live in `%APPDATA%\SpoolerCleanup`, nothing is copied anywhere).
2. **Install for** — *Just me* (no elevation, writes to `%LocalAppData%\Programs\SpoolerCleanup`,
   HKCU registry) or *All users of this PC* (UAC prompt; `%ProgramFiles%\SpoolerCleanup`,
   HKLM uninstall entry, machine-wide autostart option).
3. **Installation folder** — editable path + *Browse…* (standard folder picker).
4. **Additional options** — desktop shortcut, Start-menu shortcut, start with Windows,
   start minimized to tray at logon.
5. **Done** — the app launches the main window.

What "install" actually does (nothing else, no services, no drivers):

* copies the exe to the chosen folder;
* creates the selected `.lnk` shortcuts (Desktop / Start ▸ Programs);
* optionally adds `…\SpoolerCleanup.exe /tray` to `Run` (HKCU, or HKLM for all-users installs
  performed elevated);
* registers *Programs & Features* entry `SpoolerCleanup` with DisplayVersion, publisher,
  icon, and `UninstallString = "…\SpoolerCleanup.exe" /uninstall` (+ quiet variant).

## Silent / automated install

```cmd
SpoolerCleanup.exe /install /silent                     :: current user, defaults
SpoolerCleanup.exe /install /allusers /silent           :: triggers UAC once
SpoolerCleanup.exe /install /dir="D:\Tools\PSQC" /nodesktop /nostartmenu /noautostart /notray
```

Exit codes: `0` ok, `1` copy/IO failure, `2` elevation refused/failed.

## Portable mode

* run with `/portable`, **or** drop an empty `portable.txt` next to the exe.
* Settings file moves next to the exe (`settings.json`); the wizard is skipped;
  the main window shows a *Portable mode* chip.

## Uninstall

* Settings ▸ Data ▸ *Uninstall…*, or *Programs & Features* ▸ Uninstall, or
  `SpoolerCleanup.exe /uninstall [/silent]`.
* Removes shortcuts, autostart entry, registry entries and (after the process exits)
  the installation folder. User settings (`%APPDATA%\SpoolerCleanup`) are kept on purpose.

## Where things live

| Data | Path |
|---|---|
| Settings | `%APPDATA%\SpoolerCleanup\settings.json` (or next to exe in portable mode) |
| Logs | `%LOCALAPPDATA%\SpoolerCleanup\logs\app-YYYYMMDD.log` (14-day rotation) |
| Spool backups | `%LOCALAPPDATA%\SpoolerCleanup\backup\<timestamp>\` |
| Update download | `%TEMP%\SpoolerCleanup-update.exe` (deleted after apply) |
