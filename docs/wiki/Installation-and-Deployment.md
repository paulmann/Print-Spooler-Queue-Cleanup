# Installation and Deployment

Four deployment shapes, from "one file on one laptop" to "400 hosts on a schedule".

---

## Requirements

| | |
|---|---|
| Operating system | Windows 10 / 11, Windows Server 2012 R2 or later, with the Print Spooler feature installed |
| PowerShell | Windows PowerShell **5.1** or PowerShell **7.x** (7.x required for `-ThrottleLimit` > 1) |
| Privileges | Administrator for the local machine |
| Network | WinRM (TCP 5985/5986) plus authorization for remote targets |
| Dependencies | **None.** No modules to install, no internet access needed to run |
| Disk | A writable location for logs and backups; `%ProgramData%` by default |

---

## Option A — one file, one machine

The fastest path, and enough for a helpdesk laptop.

```powershell
$dest = "$env:USERPROFILE\Tools"
New-Item -ItemType Directory -Path $dest -Force | Out-Null

Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.ps1 -OutFile "$dest\Clear-PrintSpoolerQueue.ps1"
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/Clear-PrintSpoolerQueue.bat -OutFile "$dest\Clear-PrintSpoolerQueue.bat"

Unblock-File "$dest\Clear-PrintSpoolerQueue.ps1"
```

`Unblock-File` matters: Windows marks files from the internet with a Zone.Identifier alternate data
stream, and the default execution policy will refuse to run them.

Run it:

```powershell
cd $dest
.\Clear-PrintSpoolerQueue.ps1 -Diagnostic
.\Clear-PrintSpoolerQueue.ps1 -Force -Backup
```

Or double-click `Clear-PrintSpoolerQueue.bat` and let it elevate.

### Verifying the download

```powershell
Invoke-WebRequest https://github.com/paulmann/Print-Spooler-Queue-Cleanup/releases/latest/download/checksums.txt -OutFile checksums.txt
Get-FileHash "$dest\Clear-PrintSpoolerQueue.ps1" -Algorithm SHA256
Get-Content .\checksums.txt
```

Or have the tool check its own signature:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -VerifySignature
.\Clear-PrintSpoolerQueue.ps1 -SelfTest        # internal consistency, exits non-zero on failure
```

---

## Option B — clone the repository

For development, or when you want the source, tests and build pipeline.

```powershell
git clone https://github.com/paulmann/Print-Spooler-Queue-Cleanup.git
cd Print-Spooler-Queue-Cleanup

./tools/Build.ps1              # assembles dist/ from src/
./tests/Run-Tests.ps1          # full quality gate (425 checks)

./dist/Clear-PrintSpoolerQueue.ps1 -Diagnostic
```

`dist/` holds the generated artifacts. **Do not edit them** — `src/` is the source of truth and the
next build overwrites `dist/`.

---

## Option C — importable module

For automation that wants discoverable cmdlets rather than a script path.

```powershell
Expand-Archive .\PrintSpoolerCleanup-4.0.0.zip -DestinationPath C:\Tools\psq
Import-Module C:\Tools\psq\PrintSpoolerCleanup\PrintSpoolerCleanup.psd1

Get-Command -Module PrintSpoolerCleanup
Clear-PrintSpoolerQueue -Force -Backup
Test-PrintSpoolerHealth
```

Install for every session by copying into a module path:

```powershell
# PowerShell 7
Copy-Item .\PrintSpoolerCleanup -Destination "$env:ProgramFiles\PowerShell\Modules\" -Recurse

# Windows PowerShell 5.1
Copy-Item .\PrintSpoolerCleanup -Destination "$env:SystemRoot\System32\WindowsPowerShell\v1.0\Modules\" -Recurse
```

Then `Import-Module PrintSpoolerCleanup` from anywhere. See
[Cmdlet Reference](Cmdlet-Reference) for the 14 exported commands.

---

## Option D — network share

The classic helpdesk pattern: one copy, everyone runs it, an update is a file replacement.

```powershell
$share = '\\corp-fs01\share\Scripts\PrintCleanup'
New-Item -ItemType Directory -Path $share -Force | Out-Null
Copy-Item .\dist\Clear-PrintSpoolerQueue.ps1, .\dist\Clear-PrintSpoolerQueue.bat, `
          .\dist\PrintSpoolerCleanup.config.psd1 -Destination $share
```

Set site defaults in the config file so nobody has to remember the flags:

```powershell
# $share\PrintSpoolerCleanup.config.psd1
@{
    Cleanup   = @{ Backup = $true; BackupKeepDays = 14; OlderThan = '6h' }
    Logging   = @{ Level = 'INFO'; RetentionDays = 60 }
    Execution = @{ ExportFormat = @('Csv'); ExportPath = '\\corp-fs01\share\Logs\Spooler' }
}
```

The config is auto-discovered from the script directory, so this applies to everyone who runs the
tool from the share.

> **Read-only share?** Fine. The log target resolution skips locations it cannot write to and falls
> back to `%ProgramData%`, `%LocalAppData%` and finally `%TEMP%`. v3 wrote its log next to the script
> and silently lost all logging when the script lived somewhere read-only.

Run it:

```powershell
& '\\corp-fs01\share\Scripts\PrintCleanup\Clear-PrintSpoolerQueue.ps1' -Force
```

Note that script execution from a UNC path is subject to execution policy. Either sign the script,
set a policy that permits it for the `LocalIntranet` zone, or copy locally first.

---

## Scheduled maintenance

The installer is the recommended path — it gets the flags right so the task cannot block on a prompt.

```powershell
Copy-Item .\dist\Clear-PrintSpoolerQueue.ps1 C:\Tools\
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 `
    -Schedule Daily -StartTime 03:00 -PassThru
```

Verify it immediately rather than six months later during an incident:

```powershell
Get-SpoolerCleanupTask | Format-List TaskName, State, NextRunTime
Start-SpoolerCleanupTask -Wait          # -> ExitCode : 0
```

Full details, including event-driven self-healing on SCM 7031:
[Scheduling and Automation](Scheduling-and-Automation).

---

## Group Policy

### Startup script

```
Computer Configuration → Policies → Windows Settings → Scripts → Startup
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File \\corp\share\Scripts\Clear-PrintSpoolerQueue.ps1
      -Force -NonInteractive -Quiet -NoProgress -Backup -OlderThan 12h
```

Reading from a share means an update is a file copy.

### Scheduled task via GPO Preferences

Generate the XML once, then distribute it:

```powershell
$def = Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -PassThru
$def.Xml | Set-Content \\corp\share\GPO\PrintSpoolerCleanup.xml -Encoding Unicode
```

`Computer Configuration → Preferences → Control Panel Settings → Scheduled Tasks → New Scheduled Task
Import` and point at the XML.

The XML sets the things the GUI defaults get wrong: `DisallowStartIfOnBatteries=false`,
`StopIfGoingOnBatteries=false`, `StartWhenAvailable=true`, an execution time limit, and
restart-on-failure.

### Machine-wide policy via environment variables

`Computer Configuration → Preferences → Windows Settings → Environment`:

| Variable | Value | Effect |
|---|---|---|
| `PSQ_BACKUP_PATH` | `\\corp-fs01\share\SpoolerBackups` | Central backups — and implies `-Backup` |
| `PSQ_LOG_PATH` | `C:\ProgramData\PrintSpoolerCleanup\Logs\spooler.log` | Fixed log location |
| `PSQ_NO_UPDATE_CHECK` | `true` | Air-gapped sites |
| `PSQ_GITHUB_TOKEN` | *(user or computer secret)* | Raise the API rate limit |

Prefer a **computer** GPO for these so the values are present when a task runs as SYSTEM.

See [Configuration](Configuration) for the full list.

---

## Intune / ConfigMgr

Package as a PowerShell script or an application.

**Detection:**

```powershell
$v = & 'C:\Tools\Clear-PrintSpoolerQueue.ps1' -ShowVersion -MachineReadableJson | ConvertFrom-Json
if ($v.Version -ge [Version]'4.0.0') { Write-Output 'Installed' }
```

**Install:**

```powershell
$dest = 'C:\Tools'
New-Item -ItemType Directory -Path $dest -Force | Out-Null
Copy-Item "$PSScriptRoot\Clear-PrintSpoolerQueue.ps1" $dest -Force
Unblock-File "$dest\Clear-PrintSpoolerQueue.ps1"
& "$dest\Clear-PrintSpoolerQueue.ps1" -SelfTest -Quiet
if ($LASTEXITCODE -ne 0) { throw 'Self-test failed after deployment' }
```

`-SelfTest` as a post-install gate is worth the two seconds: it verifies the payload generates and
parses, the guards behave, the config schema validates, the help is complete and a writable log
location exists. A deployment that fails the self-test is a deployment you want to know about now,
not during the next incident.

**Uninstall:**

```powershell
& 'C:\Tools\Clear-PrintSpoolerQueue.ps1' 2>$null
Uninstall-SpoolerCleanupTask -Confirm:$false
Remove-Item 'C:\Tools\Clear-PrintSpoolerQueue.*' -Force
```

### Let the tool update itself

Once deployed, `-AutoUpdate` (included in tasks created by `Install-SpoolerCleanupTask`) checks
GitHub at most once per cooldown window and stages a new version. This removes the need to re-package
for every point release:

```powershell
Install-SpoolerCleanupTask -ScriptPath C:\Tools\Clear-PrintSpoolerQueue.ps1 -Schedule Daily -StartTime 03:00
# the task runs: ... -Force -NonInteractive -Quiet -NoProgress -Backup -AutoUpdate
```

Pin a channel to control rollout:

```powershell
# %ProgramData%\PrintSpoolerCleanup\PrintSpoolerCleanup.config.psd1
@{ Update = @{ Channel = 'stable'; CooldownHours = 168; VerifyChecksum = $true } }
```

See [Updates and Versioning](Updates-and-Versioning).

---

## Code signing

The build can sign the artifacts if you have a code-signing certificate:

```powershell
./tools/Build.ps1 -SigningThumbprint 1A2B3C4D...
```

Then require the signature at run time:

```powershell
.\Clear-PrintSpoolerQueue.ps1 -VerifySignature -ExpectedSignerThumbprint 1A2B3C4D... -Force
```

Exit code **14** on a broken, missing or unexpected signature. This is the right posture for an
environment where an execution policy of `AllSigned` is enforced, or where you want to be certain
nobody edited the copy on the share.

Signing is optional: an unsigned copy works normally, and `-VerifySignature` reports `NotSigned`
rather than failing when you have not asked for a specific publisher.

---

## Where the tool writes

| What | Default location |
|---|---|
| Log | first writable of `%ProgramData%\PrintSpoolerCleanup\Logs`, `%LocalAppData%\…`, `<script>\Logs`, `%TEMP%\…` |
| Backups | `%ProgramData%\PrintSpoolerCleanup\Backups` |
| Update state | `<script dir>\update-state.json`, else `%ProgramData%\PrintSpoolerCleanup\` |
| Staged updates | `<script dir>\.update\` |
| Version backups | `<script dir>\backup-<version>-<stamp>\` |
| Reports | `-ExportPath`, else the script directory |

`%ProgramData%` needs elevation to create the first time. If the tool runs unelevated against remote
targets only, it falls back to `%LocalAppData%` — which is why the resolution is a probe chain rather
than a fixed path.

---

## Upgrading from v3

Replace both files. There is no migration step and no state to convert.

**One breaking change to check:** the exit-code contract was extended and codes 1 and 2 swapped
meaning.

| v3 | v4 |
|---|---|
| `1` fatal error | `2` Failed |
| `2` completed with warnings | `1` PartialSuccess |

If you have monitoring rules or a wrapper script keyed on v3 codes, update them. Details and the
full 0–15 table: [Exit Codes](Exit-Codes).

Also note the batch launcher no longer implements its own cleanup — a `/UsePS` switch is not needed
because delegation is now the only behaviour.

---

## Related pages

- [Getting Started](Getting-Started)
- [Enterprise Deployment](Enterprise-Deployment)
- [Scheduling and Automation](Scheduling-and-Automation)
- [Configuration](Configuration)
- [Security and Privacy](Security-and-Privacy)
