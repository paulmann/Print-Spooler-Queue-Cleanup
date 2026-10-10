# Security Policy

## Reporting a vulnerability

**Please do not open a public GitHub issue for a security vulnerability.**

Email **Mikhail@Deynekin.com** with:

- the affected version (`.\Clear-PrintSpoolerQueue.ps1 -ShowVersion`)
- a description of the vulnerability and its realistic impact
- reproduction steps or a proof of concept
- whether you would like to be credited

You should receive an acknowledgement within 5 business days. We aim to publish a fix and a release
note within 30 days for confirmed issues of high severity, and to keep you informed throughout.

Please allow a reasonable window for a fix before public disclosure. This tool runs with
administrator privileges on print infrastructure, so a coordinated disclosure genuinely protects
people.

### Attaching diagnostics

A diagnostics bundle helps a great deal and contains **no credentials, tokens or document contents**:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic -DiagnosticsBundlePath C:\Temp
```

Review it before sending if you are unsure — `effective-config.json` redacts `Update.Token`, but you
know your environment better than the tool does.

---

## Scope

The following are in scope:

- Any path by which the tool could delete files outside a validated spool directory
- Bypass of the four spool-directory guards in `Test-SpoolDirectorySafeInternal` /
  `Test-SpoolDirSafe`
- Code execution through the configuration file (which is AST-parsed with `SafeGetValue()`
  specifically so it cannot execute anything)
- Code execution through a crafted backup archive (zip-slip in `Expand-SpoolBackupArchive`)
- Bypass of update checksum verification
- Credential, token or secret leakage into logs, reports, event records or diagnostics bundles
- Privilege escalation via `-SelfElevate` or the batch launcher's UAC path
- Injection through report output (the HTML report escapes every embedded value)

The following are **out of scope** and are documented, intended behaviour:

- The tool requires administrator privileges. There is no reduced-privilege mode, by design: a
  cleanup that cannot stop the Spooler leaves the queue worse than it found it.
- `-SkipChecksumVerification` installs unverified bytes. It is an explicit operator override,
  recorded as `VerifyResult = SkippedByOperator` and logged as a warning.
- `-PurgeDriverStore` removes unreferenced driver packages. It requires `-Repair` **and** `-Force`,
  and never removes a package that a printer references.
- Deleting `*.SPL` / `*.SHD` files from a validated spool directory is the product's purpose.
- Backups contain rendered document payloads. That is inherent to backing up a print queue. Default
  location is on-box (`%ProgramData%`); nothing leaves the machine unless the operator points
  `-BackupPath` or `-ArchivePath` at a share.
- A hostile actor who is already a local administrator on the host.

---

## Security-relevant design properties

These are the invariants that a review should focus on, and that the test suite asserts.

| Property | Where it is enforced | Verified by |
|---|---|---|
| Only `.SPL`/`.SHD` are ever deleted | Payload extension allow-list | `Unit.Guards.ps1` |
| A wildcard can never enter the allow-list | `Test-Configuration` rejects `*`, `.*`, `''` | `Unit.Configuration.ps1` |
| The spool directory is never recursed | `-File` without `-Recurse` in the payload | `Unit.Guards.ps1` (asserts no recursive op against the spool dir) |
| A UNC spool path is refused | Guard 2 | `Unit.Guards.ps1` |
| Drive-letter shape is validated **before** `GetFullPath()` | Guard 2 (platform-relative API) | `Unit.Guards.ps1` |
| A path not under a `spool` folder is refused | Guard 2 | `Unit.Guards.ps1` |
| Protected system roots are refused | Guard 3 | `Unit.Guards.ps1` |
| Absurd file counts / sizes abort | Guard 4 | `Unit.Guards.ps1` |
| Nothing is deleted if the Spooler will not stop | Payload returns `ServiceFailure` | payload review |
| Every failure path attempts to restart the Spooler | `LAST RESORT` block in the payload | `Unit.Guards.ps1` |
| Config files cannot execute code | AST parse + `SafeGetValue()` | `Unit.Configuration.ps1` (5 negative cases) |
| Downloads are SHA-256 verified | `Update-SpoolerCleanupCore` | `smoke-test.ps1` |
| A staged update is re-verified before promotion | `Invoke-PendingUpdateInternal` | code review |
| Credentials never reach a log, report or child process | `New-ChildArgumentList` excludes `Credential` | `smoke-test.ps1` |
| `Update.Token` is redacted from the bundle | `Write-DiagnosticsBundle` | code review |
| HTML report values are escaped | `ConvertTo-HtmlEncoded` on every value | `Unit.Utilities.ps1` (injection test) |
| Zip entries cannot escape the destination | `GetFileName()` on every entry name | `smoke-test.ps1` |
| The tool refuses to prompt on a non-interactive host | `Test-NonInteractiveHost` | `smoke-test.ps1` |

---

## Supported versions

| Version | Supported | Notes |
|---|---|---|
| 4.x | ✅ | Current. Receives security fixes |
| 3.1.0 | ⚠️ Best effort | Pre-guard-rail. Missing the `spool`-folder guard, the drive-letter check, backup, and selective cleanup. **Upgrade recommended** |
| ≤ 3.0 | ❌ | Contains a remote-execution bug (`Clear-SpoolFiles` referenced an undefined variable) and uses `Get-Service -ComputerName`, which does not exist in PowerShell 7 |

Run `.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate` to see where you stand. Exit code **7** means a newer
version exists.

---

## Hardening guidance

If you operate this in a sensitive environment:

1. **Sign your copy** and require the signature:
   ```powershell
   ./tools/Build.ps1 -SigningThumbprint <thumbprint>
   .\Clear-PrintSpoolerQueue.ps1 -VerifySignature -ExpectedSignerThumbprint <thumbprint>
   ```
2. **Pin the update channel** to `stable`, or disable updates and deploy through your own pipeline:
   ```powershell
   @{ Update = @{ Enabled = $false; AutoCheck = $false } }
   ```
3. **Keep `Update.VerifyChecksum = $true`** (the default). Never set
   `-SkipChecksumVerification` in automation.
4. **Use `PSQ_GITHUB_TOKEN` from a secrets manager**, not `-GitHubToken` on a command line.
5. **Centralise backups to a share with document-appropriate ACLs** — spool files contain rendered
   documents, so that share inherits the sensitivity of the documents themselves.
6. **Audit with the Event Log**: `Logging.EventLog = $true` gives one SIEM query for the whole site.
7. **Run `-SelfTest` after every deployment** so a corrupted copy is caught before an incident:
   ```powershell
   .\Clear-PrintSpoolerQueue.ps1 -SelfTest -Quiet; if ($LASTEXITCODE) { throw }
   ```

---

## Disclosure of known limitations

Stated plainly, because a runbook that hides them is worse than one that lists them:

- **Restoring spool files does not re-queue jobs.** `spoolsv.exe` builds its job table in memory and
  only reads files it created in the current boot. A restore recovers the document payload for
  re-submission; it is not a rollback.
- **`.SHD` metadata is best-effort.** The header layout changed across Windows versions, so when WMI
  is unavailable the printer and document names are extracted heuristically. Jobs with no readable
  metadata are *not* selected by name-based filters, and the run reports this rather than guessing.
- **A forced service stop terminates `spoolsv.exe`.** `Service.ForceStopFallback` (default on) only
  fires after the normal stop has already timed out. The SCM treats an exited service process as
  stopped and applies its recovery policy, but any in-flight job on that host is lost.
- **The batch launcher cannot update itself while running.** `cmd.exe` re-reads a batch file by byte
  offset during execution. The launcher's own staged update lands on the following start, and it says
  so.

---

## Contact

**Mikhail Deynekin** — <Mikhail@Deynekin.com> — [Deynekin.com](https://Deynekin.com)

Repository: <https://github.com/paulmann/Print-Spooler-Queue-Cleanup>
