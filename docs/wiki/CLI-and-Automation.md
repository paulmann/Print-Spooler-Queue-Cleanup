# CLI and automation (v5)

The same `SpoolerCleanup.exe` serves GUI and automation. CLI modes attach to the parent
console when present (`cmd`, PowerShell, CI runners).

## Commands

| Command | Effect | Exit code |
|---|---|---|
| `/help`, `/?`, `-h` | usage text | 0 |
| `/status` | `spooler=<running|stopped|pending|unknown> jobs=N printers_with_jobs=M version=…` | 0 running, 1 otherwise |
| `/list` | per-printer job table (`#id document owner size status`) or `(empty)` | 0 |
| `/purge` | full guarded purge (backup per settings), prints `OK: removed N job(s)[; backup: path]` or `FAIL: reason (code N)` | 0 ok, 1 generic, 2 elevation required |
| `/install […]` | silent install (see INSTALL.md) | 0/1/2 |
| `/uninstall [/silent]` | remove installation | 0 |
| `/tray`, `/minimized` | start hidden to tray | — |
| `/portable` | settings next to exe, skip wizard | — |
| `/lang=<tag>` | force UI language for this session (and persist) | — |
| `/theme=<system|light|dark>` | force theme (persist) | — |

Unknown switches are ignored (forward compatibility).

## Examples

```cmd
rem CI: fail the job if the queue is not empty after the test
SpoolerCleanup.exe /status || echo queue not empty

rem Support script: clear and, if needed, elevate once
SpoolerCleanup.exe /purge
if errorlevel 2 powershell -c "Start-Process SpoolerCleanup.exe -Verb RunAs -ArgumentList '/purge' -Wait"

rem Scheduled task (Task Scheduler): daily 03:00 silent purge
schtasks /create /tn PSQCCleanup /tr "\"C:\Tools\SpoolerCleanup.exe\" /purge" /sc daily /st 03:00
```

## Exit code contract

`0` success · `1` generic failure (see message) · `2` elevation required/refused ·
`3` self-update replace scheduled for reboot (`/replace` child only).
