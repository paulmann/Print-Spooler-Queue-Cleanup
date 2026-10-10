# Backup and Restore

`-Backup` archives every file the tool is about to delete, before it deletes anything. On a print
server whose queues carry invoices, prescriptions, legal filings or payslips, this is the difference
between a fix and an incident.

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup
```

---

## What a backup contains

A single ZIP named `spool-backup_<computer>_<yyyyMMdd-HHmmss>.zip`:

```
spool-backup_SRV01_20261010-031500.zip
├── backup-manifest.json
├── F00412.SHD
├── 00412.SPL
├── F00413.SHD
└── 00413.SPL
```

The manifest records:

| Field | Purpose |
|---|---|
| `Tool`, `ToolVersion` | Which build created the archive |
| `CreatedUtc` | When |
| `ComputerName` | Which host |
| `SpoolDirectory` | The exact path the files came from |
| `FileCount`, `Files[]` | Every file with its name and size |
| `Jobs[]` | Job id, printer, document name, owner, size, submission time |
| `RestoreNote` | The honest limitation, stated in the artifact itself |

The run report records `BackupPath` and `BackupHash` (SHA-256) per computer, so an archive can be
verified later against the report that created it.

**Only the files that were selected are archived.** If you ran with `-OlderThan 2h`, the healthy jobs
are not in the archive because they were never touched.

---

## Where backups go

| Setting | Location |
|---|---|
| Default | `%ProgramData%\PrintSpoolerCleanup\Backups` |
| `-BackupPath <dir>` | Per-run override |
| `Cleanup.BackupPath` in the config file | Site-wide default |
| `PSQ_BACKUP_PATH` environment variable | Machine-wide policy (also implies `-Backup`) |

Move archives to a share afterwards with `-ArchivePath`:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -ArchivePath \\filesrv\logs\spooler-backups
```

This matters when the spool volume is the thing that is full (root cause 6) — point `-BackupPath` at
a *different* volume, or the backup competes with the problem it is meant to protect against.

---

## Retention

`Cleanup.BackupKeepDays` (default 30) prunes bundles older than the window on every run. Pruning is
scoped strictly to the backup directory and to the `spool-backup_*.zip` prefix, so a mis-pointed
`-BackupPath` can never turn retention into a recursive delete.

Set it to `0` to disable pruning — appropriate where backups are subject to a retention policy you
manage elsewhere.

List what is on hand:

```powershell
Get-SpoolerBackup | Format-Table CreatedUtc, ComputerName, FileCount, SizeBytes, Valid
Get-SpoolerBackup -BackupPath \\filesrv\logs\spooler-backups
```

---

## Requiring a backup

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup -BackupRequired
```

With `-BackupRequired`, a failed backup is a **hard abort**: nothing is deleted and the run fails.
Without it, a backup failure is a warning and the cleanup proceeds.

Use `-BackupRequired` wherever the queue has legal, financial or clinical significance. The default
is deliberately permissive, because a tool that refuses to clear a wedged queue when its backup share
is offline creates a worse outage than the one it was called to fix.

---

## Restoring

```powershell
Restore-SpoolerBackup -Path .\spool-backup_SRV01_20261010-031500.zip

Restore-SpoolerBackup -Path .\spool-backup_SRV01_20261010-031500.zip `
                      -DestinationDirectory C:\Temp\restored -Overwrite
```

The archive is validated and its SHA-256 computed before anything is extracted. Extraction is
zip-slip-safe: entry names are reduced to their file-name component, so an archive carrying
`..\..\` segments cannot write outside the destination.

Verify against the hash recorded in the original run report:

```powershell
$run = Import-Clixml .\run.xml      # or read the JSON report
Restore-SpoolerBackup -Path $run.Results[0].BackupPath -ExpectedHash $run.Results[0].BackupHash
```

A mismatch throws and nothing is extracted.

---

## ⚠️ The honest limitation

> **Copying spool files back into `PRINTERS` does NOT re-queue the job.**

`spoolsv.exe` maintains its job table **in memory**. The `.SHD`/`.SPL` files are how that table
survives a restart, but on start-up the Spooler only re-materialises jobs it recognises as its own
from the current boot sequence. Files dropped into the folder from outside are ignored, or garbage
that is eventually reaped.

So a restore is a **forensic** recovery, not a transactional rollback:

| What a restore gives you | What it does not give you |
|---|---|
| The rendered document payload (`.SPL`) | A job back in the queue |
| The job metadata (`.SHD`, manifest) | The user's print settings reapplied automatically |
| Evidence of what was deleted | An undo button |

**To actually re-print a recovered document**, re-submit it from the application. The `.SPL` contains
EMF or RAW data; for RAW (most network printers) it can often be sent straight to the device:

```powershell
# RAW spool data -> printer, bypassing the queue entirely
Copy-Item .\restored\00412.SPL .\recovered-job.prn
# then send it with your usual RAW print utility, or:
Get-Content .\recovered-job.prn -Raw -Encoding Byte | Out-Printer -Name "<printer>"
```

For EMF spool data this does not work directly — the document must be re-rendered from the original
application. The manifest's `DocumentName` and `Owner` fields tell you who to ask.

This limitation is stated in four places on purpose: in `Restore-SpoolerBackup`'s output, in the
manifest's `RestoreNote`, in the HTML report under the backup table, and here. A runbook that
promises a rollback that cannot happen gets someone paged at three in the morning.

---

## Restore failures

`Restore-SpoolerBackup` throws, and the script exits **13** (`RestoreFailed`), when:

- the archive is missing or unreadable
- no `backup-manifest.json` is inside it (it is not one of ours)
- `-ExpectedHash` does not match
- the destination cannot be created or written

Inspect an archive without extracting it:

```powershell
Get-SpoolBackupInfo -Path .\spool-backup_SRV01_20261010-031500.zip | Format-List *
```

---

## Backups across a fleet

Each target produces its own archive **on that target**, because the payload runs remotely and the
files never traverse the network twice. Collect them centrally with `-ArchivePath` pointing at a
share:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -ComputerName (Get-Content .\hosts.txt) -ThrottleLimit 16 `
    -Force -Backup -ArchivePath \\filesrv\logs\spooler-backups
```

The report lists every archive with its SHA-256, so you can reconcile what was created against what
arrived:

```powershell
$run = .\Clear-PrintSpoolerQueue.ps1 -ComputerName $hosts -Force -Backup -PassThru
$run.Results | Where-Object BackupPath |
    Select-Object ComputerName, BackupPath, BackupHash | Export-Csv backups.csv -NoTypeInformation
```

---

## Related pages

- [Safety Model](Safety-Model)
- [Reports and Machine-Readable Output](Reports-and-Machine-Readable-Output)
- [Print Spooler Internals](Print-Spooler-Internals#the-job-table-is-in-memory) — why the limitation exists
- [Enterprise Deployment](Enterprise-Deployment)
