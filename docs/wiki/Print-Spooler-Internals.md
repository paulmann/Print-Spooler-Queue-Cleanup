# Print Spooler Internals

You do not need this page to use the tool. You do need it to understand *why* the tool is shaped
the way it is, and to diagnose the cases where clearing the queue is not enough.

---

## The service

`Spooler` (`spoolsv.exe`) runs as `LocalSystem` and owns the print pipeline:

```
Application
   |  OpenPrinter / StartDocPrinter / WritePrinter / EndDocPrinter
   v
GDI / XPS print path  (splwow64.exe for 32-bit apps on 64-bit Windows)
   |
   v
spoolsv.exe  ----> local port monitor  ----> USB / LPT
              \---> network port monitor ----> TCP 9100 / IPP / WSD / SMB
```

Two process identities are involved, and confusing them explains a lot of behaviour:

- **`spoolsv.exe`** — the service itself. Owns the job table, the spool directory and the port
  monitors.
- **`splwow64.exe`** — the *WOW* print provider, spawned per-session to host 32-bit print drivers on
  64-bit Windows. A wedged `splwow64.exe` can hold handles open and keep the Spooler in
  `STOP_PENDING`, which is why the tool's force-stop fallback exists.

---

## The spool directory

Default: `%SystemRoot%\System32\spool\PRINTERS`

Overridable via the registry:

```
HKLM\SYSTEM\CurrentControlSet\Control\Print\Printers
    DefaultSpoolDirectory  (REG_SZ)
```

Each queued job is a **pair** of files sharing a five-digit number:

| File | Contents |
|---|---|
| `F#####.SHD` | **Sh**adow file — job metadata: submitting user, machine, document name, printer, priority, page range, form, deviation/duplex settings. Small. |
| `#####.SPL` | **Spool** file — the rendered print data in EMF or RAW form. This is where the gigabytes are. |

The `F` prefix on the shadow file is the only naming difference; the numeric part is the job's
spool number.

Two consequences matter operationally:

1. **Deleting only the `.SPL` orphans the job.** The `.SHD` is the source of truth for the job
   table; a lone `.SPL` is garbage the Spooler will eventually reap. A tool that deletes one and not
   the other leaves a queue that looks empty but is not. This tool always deletes the pair.
2. **Deleting only the `.SHD` removes the job but leaves the data behind.** This is why a selective
   delete must resolve *both* file names for each selected job, which is what the payload's
   `Get-FilesForJob` does.

Other files you may see in the same directory:

| Pattern | Meaning |
|---|---|
| `~DF*.tmp` | Transient files created while the Spooler is writing a job |
| `*.SPF` | Legacy spool format left by some older drivers |
| `*.BBL`, `*.EMF`, driver-specific | Vendor artefacts |

By default the tool touches **only** `.SPL` and `.SHD`. `Cleanup.PurgeLegacyExtensions = $true`
adds `.SPF` and `~DF*`. It never deletes anything else, never recurses, and refuses a configuration
that tries to introduce a wildcard.

---

## The job table is in memory

This is the single most misunderstood fact about the Print Spooler, and the reason "restore" cannot
mean "rollback".

`spoolsv.exe` maintains its job table **in memory**. The `.SHD`/`.SPL` files are how that table
survives a restart, but on start-up the Spooler only re-materialises jobs it recognises as its own
from the current boot sequence. Copying files back into `PRINTERS\` does not re-queue them.

So a `-Backup` archive is a **forensic** copy: it recovers the rendered document payload so an
operator can re-submit it from the application, or hand the `.SPL` to a user who needed that
invoice. It is not a transactional undo, and every place the tool mentions restore says so
explicitly — in `Restore-SpoolerBackup`'s output, in the manifest's `RestoreNote`, and in the HTML
report.

---

## Dependent services

`Get-Service Spooler | Select-Object -ExpandProperty DependentServices` on a typical Windows 11 or
Server host returns some subset of:

| Service | Display name | Why it matters |
|---|---|---|
| `PrintNotify` | Printer Extensions and Notifications | **The usual culprit.** Holds handles into the spooler; a stop that does not take `PrintNotify` down first commonly hangs in `STOP_PENDING` |
| `Fax` | Fax | Depends on the Spooler for its transport |
| `PrintWorkflowUserSvc_<hash>` | Print Workflow User Service | Per-user, dynamically named. Handles XPS/IPP print workflow. There can be several |
| `WSDPrintDevice` (varies) | Web Services print support | Vendor monitor helpers |

The tool stops these **before** the Spooler and restarts them **after** it — and only restarts the
ones that were actually `Running`. A `Fax` service that was disabled stays disabled; a cleanup must
never silently change a host's configuration.

Because `PrintWorkflowUserSvc` instances carry a per-user suffix, the tool matches them by prefix
rather than by exact name.

---

## Querying the queue

Two sources, with different failure characteristics:

### WMI — `Win32_PrintJob`

```powershell
Get-CimInstance Win32_PrintJob | Select-Object Name, Document, Owner, Size, TotalPages, JobStatus, TimeSubmitted
```

`Name` is `"<printer name>, <job id>"`. Rich metadata, authoritative — but it is served **by the
Spooler**, so when the Spooler is hung the query hangs or returns nothing. That is exactly the
situation in which you are running a cleanup tool.

### The files themselves — `.SHD` headers

Always available, because it is just a file read. The `.SHD` header layout changed across Windows
versions, so rather than marshalling a version-specific struct the tool scans the first 256 KB for
the two shapes that are stable in practice:

- the UNC printer path `\\<machine>\<printer>`
- the source document path `<drive>:\<...>\<file.ext>`

Both appear as UTF-16LE (Unicode SHD) or ANSI (legacy SHD). The job id comes from the file name.

This is **best-effort**: it reliably yields a job id, size and timestamp, and usually a printer and
document name, but it is not guaranteed. That is why the tool:

1. prefers WMI while the Spooler answers;
2. correlates WMI results with the `.SHD` scan (positional pairing is reliable because the Spooler
   assigns file numbers monotonically in submission order and `Win32_PrintJob` enumerates in the
   same order) so a job gets *both* rich metadata and exact file names;
3. reports unmatched `.SHD` pairs as **orphans** rather than guessing;
4. records `Source` on every job — `WMI`, `WMI+SHD` or `SHD` — so you always know how much to trust
   the metadata;
5. **refuses to apply name-based filters to jobs with no readable metadata**, and says so in the
   warnings. Silently guessing which printer a job belongs to and then deleting it would be worse
   than doing nothing.

---

## Why `net stop spooler` hangs

The sequence, and where it goes wrong:

1. The SCM sends `SERVICE_CONTROL_STOP` to `spoolsv.exe`.
2. `spoolsv.exe` tries to flush and close every open job.
3. A driver call is blocked inside a port monitor waiting on a network printer that is not answering,
   or `PrintNotify` holds a handle, or `splwow64.exe` is wedged.
4. The service reports `STOP_PENDING`. The SCM waits `WaitToKillServiceTimeout` (default 20 s in
   `HKLM\SYSTEM\CurrentControlSet\Control`), then gives up.
5. `net stop` returns an error. The service is in an undefined state, the files are still locked,
   and the operator now has a print server that is neither running nor stopped.

What actually works:

- take the dependents down **first**, so nothing is holding a handle
- poll the service state rather than trusting the cmdlet's return code
- allow escalation to terminating the service process, which the SCM interprets as a clean stop
- if the stop still fails, **do not delete anything** — return `ServiceFailure` (exit 12) and leave
  the service exactly as it was found

That is precisely the sequence the payload implements, and every transition is recorded in
`serviceTrace[]` so you can see which step took how long.

---

## Print isolation

Since Windows 8 / Server 2012, drivers can run in **isolation mode 2** (fully isolated), in a
separate process from `spoolsv.exe`. Check it with:

```powershell
Get-PrinterDriver | Select-Object Name, IsKernelMode, PrinterEnvironment
```

Isolation is why a crashing driver no longer takes the whole Spooler down — and why a hung isolated
driver can leave the Spooler healthy but a single queue permanently stuck. In that case clearing the
queue works, and `-PurgeDriverStore` can remove the offending package if nothing references it any
more.

---

## Event logs worth knowing

| Log | Provider / IDs | What it tells you |
|---|---|---|
| `System` | Service Control Manager **7031** | A service terminated unexpectedly and was restarted — the classic spooler crash |
| `System` | Service Control Manager **7034** | A service terminated unexpectedly and was *not* restarted |
| `System` | Service Control Manager **7000/7009** | Failed to start / start timeout |
| `Microsoft-Windows-PrintService/Admin` | 307, 315, 316, 842, 848 | Job lifecycle, driver blocking, rendering failures |
| `Microsoft-Windows-PrintService/Operational` | everything | Verbose; **disabled by default** — enable it before a deep investigation |

`-Diagnostic` reads the first two automatically (`Diagnostics.EventLogHours`, default 72 h).

Enable the verbose log when you are chasing an intermittent driver problem:

```powershell
wevtutil sl Microsoft-Windows-PrintService/Operational /e:true
```

An `OnEvent` scheduled task triggered on SCM 7031 turns a spooler crash into a self-healing event
rather than a user ticket — see
[Scheduling and Automation](Scheduling-and-Automation).

---

## Related pages

- [Why Print Queues Get Stuck](Why-Print-Queues-Get-Stuck) — the seven root causes
- [Diagnostics Guide](Diagnostics-Guide) — reading `-Diagnostic` output
- [Safety Model](Safety-Model) — what the tool will and will not delete
- [Repair and Deep Cleaning](Repair-and-Deep-Cleaning)
