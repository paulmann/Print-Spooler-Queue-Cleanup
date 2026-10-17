# print-spooler-queue-cleanup

**Fix a stuck Windows print queue from the terminal** - see every queued job, delete the one
that wedges the queue, clear everything, pause or restart the Print Spooler. One
self-contained `.exe` inside the package: no .NET, no VC++ redistributable, no install
script, nothing downloaded at install time.

Same binary as the [GitHub release](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases).

## Why

A print job that fails to render **wedges the whole queue**: the spooler still reports
*Running*, the job sits in *Deleting…* and everything behind it stalls. Cancelling does
nothing, restarting the Print Spooler service throws away every healthy document too, and
hand-deleting `C:\Windows\System32\spool\PRINTERS` needs an elevated console, a stopped
service and a steady hand.

This tool shows the queue and removes exactly what is stuck.

## Install

```cmd
npx print-spooler-queue-cleanup --help          :: run without installing
npm install -g print-spooler-queue-cleanup      :: or install the command
```

Requires Windows 10/11 x64. Node.js is needed to install and to run the wrapper; the
executable itself has no runtime dependencies.

## Usage

```
spooler-cleanup                     open the app (setup wizard on first run)
spooler-cleanup /list               print the job table
spooler-cleanup /status             spooler state + counters
spooler-cleanup /purge              clear all queues and exit (0 ok, 2 = needs admin)
spooler-cleanup /tray               start minimized to the tray
spooler-cleanup /portable           keep settings next to the exe
spooler-cleanup /theme=dark         light | dark | system
spooler-cleanup /lang=ru            any of the 30 UI languages
spooler-cleanup /install /allusers /silent
spooler-cleanup /uninstall /silent
spooler-cleanup --version           wrapper + binary version
```

`--help` and `-h` are accepted as aliases of `/help`; everything else is passed to the
executable unchanged, and its exit code becomes the exit code of the command.

## Example

```cmd
> spooler-cleanup /status
spooler=running jobs=0 printers_with_jobs=0 version=5.1.1

> spooler-cleanup /list | findstr /i "stuck"
```

Piping works: output is UTF-8 when stdout is a pipe or a file and uses the console codepage
otherwise.

## What it will not do

- It does not touch files other than `*.SPL` / `*.SHD` in the spool folder.
- The optional backup of spool files before a bulk purge is on by default; the restore
  dialog puts them back.
- It only asks for elevation when an operation needs it.

## License

[PolyForm Noncommercial License 1.0.0](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE) -
noncommercial use only. The full text ships in the package as `LICENSE`.

Source, issues and the full documentation:
[github.com/paulmann/Print-Spooler-Queue-Cleanup](https://github.com/paulmann/Print-Spooler-Queue-Cleanup).
