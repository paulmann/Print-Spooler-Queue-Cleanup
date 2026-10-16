# Installer, uninstall and updates (v5)

## Built-in installer (no external engine)

First run → wizard: *Install* (scope → folder → options → done) or *Run once without
installing*. Scope **Just me**: no UAC, `%LocalAppData%\Programs\SpoolerCleanup`, HKCU
uninstall/autostart keys. Scope **All users**: one UAC prompt (the install runs in an
elevated child with `/install /allusers /silent`), `%ProgramFiles%\SpoolerCleanup`, HKLM
uninstall key, machine autostart when elevated.

Installed artifacts: exe copy, optional Desktop + Start-menu `.lnk`, optional `Run`
autostart value (`"…\SpoolerCleanup.exe" /tray`), Programs & Features entry with quiet
uninstall string. Nothing else: no services, no drivers, no shell extensions.

## Uninstall

GUI (Settings ▸ Data ▸ Uninstall…), Apps & Features, or `/uninstall [/silent]`.
Removes shortcuts, autostart, registry entries and the install folder (deleted by a
detached `cmd` after process exit). User settings are intentionally kept.

## Updates

* Check: `GET https://api.github.com/repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest`
  (WinHTTP, TLS, 10–30 s timeouts); asset selection prefers `*x64*.exe`.
* Trigger: About/tray button, or silent daily check when enabled; `skipVersion` honors
  "not this version".
* Apply: download to `%TEMP%\SpoolerCleanup-update.exe` → size sanity → detached child
  `/replace <pid> <target> <downloaded>` waits for the parent, swaps files atomically
  (`MoveFileEx`), relaunches; locked-file fallback schedules the replace at reboot
  (exit code 3) and tells the user.
* Store builds (`PSQC_STORE_BUILD`) disable the in-app updater — the Store updates them.

## Integrity & safety

* Single-instance mutex: a second launch activates the running window instead of
  double-purging.
* All file operations stay inside documented folders (install dir, `%APPDATA%`,
  `%LOCALAPPDATA%`, `%TEMP%` for the update blob, spool dir for `*.SPL`/`*.SHD` only).
* Every elevated path is an explicit child process with a UAC prompt; the GUI itself
  always runs `asInvoker`.
