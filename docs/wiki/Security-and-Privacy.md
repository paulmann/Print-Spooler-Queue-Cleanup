# Security and Privacy

This tool runs as an administrator, deletes files, and can download and install code from the
internet. Each of those is a legitimate thing for a print-queue maintenance tool to do, and each is a
reason to be careful. This page describes what the tool actually does and what it refuses to do.

---

## Threat model

| Asset | Threat | Mitigation |
|---|---|---|
| Files outside the spool directory | A corrupt registry value or a typo points the delete step elsewhere | Four independent guards; any failure aborts before a single delete |
| The queue itself | Deleting healthy jobs along with the wedged one | Selection filters; `-WhatIf`; `-Backup`; the plan states ALL vs SELECTED before you confirm |
| Print availability | A failed cleanup leaves the Spooler stopped | Every failure path makes a last-resort restart attempt |
| The tool binary | A tampered copy on a share, or a malicious update | Authenticode verification, SHA-256 on every download, staged self-update |
| Credentials | Leaking into a log, a report, or a child process | Never written anywhere; never forwarded to an elevated child |
| The config file | A config from a share that executes code | AST parsing with `SafeGetValue()`; commands, pipelines and assignments rejected |
| Document contents | Spool files contain other people's documents | Backups stay on-box by default; diagnostics bundles never contain document data |
| The report | Stored XSS in a ticketing system | Every HTML value is escaped; no script, no external resources |

---

## Privileges

Administrator rights are required for the local machine. There is no reduced-privilege mode.

That is deliberate rather than an oversight: a cleanup that cannot stop the Spooler deletes files that
`spoolsv.exe` still has open, which either fails or leaves the queue in a worse state than it started.
A partial fix here is worse than no fix.

Without elevation:

- an interactive run shows a banner and exits **5** (`NotAdministrator`)
- `-SelfElevate` triggers a UAC prompt and re-launches
- a non-interactive run (Task Scheduler, CI) fails immediately with exit 5 rather than hanging

The check uses `WindowsPrincipal.IsInRole(Administrator)`, not the `net session` trick, which gives
false negatives on domain controllers and on hosts where the Server service is stopped.

For remote targets the caller needs local admin on each, or you pass `-Credential`.

---

## Credentials and secrets

| Secret | Handling |
|---|---|
| `-Credential` (`PSCredential`) | Held in memory for the process lifetime. **Never** written to the log, the report, or a diagnostics bundle. **Never** forwarded to an elevated child process — `New-ChildArgumentList` excludes it explicitly |
| `-GitHubToken` | Used only in an `Authorization` header. Never logged |
| `PSQ_GITHUB_TOKEN` | Same. Preferred over `-GitHubToken` so the value never appears in a command line or a process listing |
| `Update.Token` in a config file | Redacted to `<redacted>` in the diagnostics bundle's `effective-config.json` |

Prefer the environment variable over both the parameter and the config file:

```powershell
$env:PSQ_GITHUB_TOKEN = 'ghp_...'      # set by your secrets manager, not typed
```

A token on a command line is visible in `Get-CimInstance Win32_Process`, in scheduled-task
definitions, and in shell history.

---

## Supply chain: updates

The updater is the highest-risk feature in the product, because it downloads and installs code. The
controls, in order:

### 1. Verification is mandatory by default

Every download is checked against the SHA-256 published in the release's `checksums.txt`.

| Situation | Behaviour |
|---|---|
| Checksum matches | Install, record `VerifyResult = ChecksumMatch` |
| Checksum mismatches | **Discard the download.** The installed copy is untouched. Exit 9 |
| No checksums published | **Refuse to install** unless `-SkipChecksumVerification` is passed explicitly |
| `-SkipChecksumVerification` | Install, record `VerifyResult = SkippedByOperator`, log a warning |

`Update.VerifyChecksum = $false` in a config file is possible and is a bad idea; the wiki says so in
the schema table.

### 2. The downloaded script is parsed before it is accepted

A truncated or corrupt download has syntax errors. The updater parses it and discards it on failure,
which catches the case where a proxy returns an HTML error page with a 200 status.

### 3. A process never overwrites the file it is executing

Updates are **staged** into `<install>\.update\` with a `pending-update.json` manifest recording each
file's size and SHA-256. The swap happens at the *start* of the next run.

The staged files are **re-verified against the manifest** when the swap happens, so a `.update`
directory tampered with between runs cannot be promoted. Files are moved, not copied, so a
half-applied update cannot be applied twice.

### 4. The previous version is always kept

`<install>\backup-<version>-<stamp>\` — rollback is a file copy.

### 5. Discovery has three independent strategies

`/releases/latest` → `/tags` → the raw branch file. If the API is rate-limited or blocked, discovery
still works, and the fallback is recorded in `Source` so you know which path was used.

### Pinning

For a controlled environment, disable automatic updates entirely and deploy by your own pipeline:

```powershell
@{ Update = @{ Enabled = $false; AutoCheck = $false } }
```

or `$env:PSQ_NO_UPDATE_CHECK = 'true'`.

To follow only reviewed releases:

```powershell
@{ Update = @{ Channel = 'stable'; IncludePrerelease = $false; CooldownHours = 168 } }
```

To use an internal mirror, set `Update.Repository` and `Update.Branch` — and mirror `checksums.txt`
with the files, or clients will refuse to install.

---

## Code signing

```powershell
# Build signed artifacts (maintainer)
./tools/Build.ps1 -SigningThumbprint 1A2B3C4D...

# Require a specific publisher at run time
.\Clear-PrintSpoolerQueue.ps1 -VerifySignature -ExpectedSignerThumbprint 1A2B3C4D... -Force
```

`Test-ScriptSignature` distinguishes the cases that matter:

| Status | Meaning |
|---|---|
| `Valid` | Signed, chain trusted, and the expected thumbprint/subject matches |
| `NotSigned` | No signature — normal for an open-source tool, **not** the same as broken |
| `Invalid` | Signature present but untrusted, or the file hash does not match |
| `ThumbprintMismatch` | Valid signature from an unexpected publisher |
| `NotSupported` | The platform cannot verify Authenticode |

Exit code **14** on anything other than `Valid`.

This matters most for a tool distributed from a network share that many administrators can run: sign
it, then require the signature, and an edited copy on the share stops working.

`-SelfTest` and `-ShowVersion -Detailed` report the signature status of the installed files, and the
diagnostics bundle includes `file-integrity.json` with size, SHA-256, last-write time and signature
status — which is how you confirm after an incident that the copy that ran was the copy you shipped.

---

## Configuration files are data, never code

A config file arrives from a network share, a Git repository, a deployment package. Treating it as
code would make it a remote-code-execution vector with a friendly name.

`PrintSpoolerCleanup.config.psd1` is parsed with the **PowerShell AST** and evaluated with
`SafeGetValue()`. Rejected outright, with the offending line named:

```powershell
@{ Service = @{ TimeoutSeconds = (Invoke-Expression '1+1') } }   # command invocation
@{ Service = @{ Name = 'Spooler' } } | Out-Null                  # pipeline
$x = 1; @{ Service = @{ TimeoutSeconds = $x } }                  # assignment
Remove-Item C:\ -Recurse -Force                                  # any statement
```

Files larger than 1 MB are refused. A malformed file reports a parse error rather than throwing.

Validation is strict and loud: unknown sections, unknown keys, non-numeric values in numeric fields,
out-of-range numbers, invalid log levels, invalid export formats and invalid update channels all
abort with exit **3**. A silently ignored typo is how a site discovers six months later that its
backups were never on.

Two values are refused regardless of range, because they would defeat the delete guards:

- `Cleanup.IncludeExtensions` may never contain `'*'`, `'.*'` or an empty string
- `Cleanup.Recurse` must remain `$false`

---

## Privacy

### What is recorded

Operational metadata only: computer names, service states, file names and sizes, job ids, printer
names, document names, submitting user names, timestamps, counters, warnings, errors.

### What is never recorded

- Credentials, tokens or any secret
- Document **contents**
- Passwords, hashes or authentication material
- Anything from outside the print subsystem

### Spool files contain real documents

This is the one genuine privacy consideration, and it is inherent to the problem rather than a design
choice: a `.SPL` file *is* someone's rendered invoice, prescription or payslip.

Consequences worth knowing:

- `-Backup` copies those files into an archive. The default location is `%ProgramData%` on the target
  host — **not** a network share. Nothing leaves the machine unless you point `-BackupPath` or
  `-ArchivePath` at a share.
- If you do centralise backups, that share now holds document payloads and needs the access controls
  and retention policy of the documents themselves, not of a log directory.
- `Cleanup.BackupKeepDays` (default 30) bounds how long they persist. Set it lower where document
  retention is sensitive, or to `0` and manage pruning yourself.
- The diagnostics bundle contains the job **metadata** (names, owners, sizes) but never file contents.

### Reports

Job metadata — including `DocumentName` and `UserName` — appears in the JSON, CSV, HTML and Text
reports. That is what makes them useful for finding who submitted the wedged job, and it also means
a report attached to a public issue discloses document names. Use `-DiagnosticsBundlePath` for issue
reports: it carries the metadata you need and redacts secrets, and you can review it before uploading.

---

## Network access

The tool makes **no** network calls during a cleanup. Network access happens only for:

| Feature | Destination | When |
|---|---|---|
| Remote cleanup | WinRM 5985/5986 on the targets you name | Only with `-ComputerName` |
| Update check | `api.github.com`, `raw.githubusercontent.com` | Only with `-CheckUpdate`, `-Update`, or `-AutoUpdate` (cooldown-limited) |
| Update download | `objects.githubusercontent.com` (release assets) | Only with `-Update` |

`-AutoUpdate` is off unless you ask for it, except in tasks created by `Install-SpoolerCleanupTask`
(where it is on by default and can be disabled with `-NoAutoUpdate`). In an air-gapped environment:

```powershell
@{ Update = @{ Enabled = $false; AutoCheck = $false } }
```

Outbound proxy configuration is honoured through the standard `HTTPS_PROXY` environment variable.

---

## Denial-of-service considerations

The tool stops the Print Spooler. On a shared print server that means printing is unavailable for the
10–20 seconds the operation takes.

Controls that limit the blast radius:

| Control | Effect |
|---|---|
| `-ThrottleLimit` default **1** | Sequential by default, so a fleet run cannot take down a site at once |
| `-OlderThan` | Healthy jobs are left alone |
| `-PrinterName` | Only the affected queue is touched |
| `-WhatIf` | Full plan, zero changes, service not even stopped |
| `-Watchdog` | A wedged run is terminated and the Spooler restarted |
| `Execution.ConfirmTimeoutSeconds` | An unanswered prompt cannot block a maintenance window |
| `WatchdogMinutes` / task `ExecutionTimeLimit` | A hung run cannot consume the host indefinitely |

Raise `-ThrottleLimit` knowingly, off-hours, or grouped by AD site.

---

## Reporting a security issue

Please do **not** open a public issue for a vulnerability.

Email **Mikhail@Deynekin.com** with:

- the version (`.\Clear-PrintSpoolerQueue.ps1 -ShowVersion`)
- a description of the issue and its impact
- reproduction steps or a proof of concept
- whether you want to be credited

A diagnostics bundle helps, and it never contains credentials:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

See [SECURITY.md](https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/SECURITY.md) for
the disclosure policy and supported versions.

---

## Related pages

- [Safety Model](Safety-Model) — the delete guards in detail
- [Updates and Versioning](Updates-and-Versioning)
- [Configuration](Configuration)
- [Backup and Restore](Backup-and-Restore)
- [For Developers](For-Developers) — the build and test pipeline
