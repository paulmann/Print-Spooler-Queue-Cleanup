# Updates and Versioning

The tool can check for, download, verify and install its own updates. The design is dominated by one
rule: **never brick a working tool.** A print server that cannot run the cleanup script is worse off
than one running a slightly old version of it.

---

## Checking

```powershell
.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate
Get-SpoolerCleanupUpdate                    # cmdlet form
Get-SpoolerCleanupUpdate -Channel stable
```

```
 ==============================================================================
  UPDATE CHECK
 ==============================================================================
    Installed         : 4.0.0
    Latest published  : 4.1.0
    Discovery source  : releases/latest
    Update available  : True
    Tag               : v4.1.0
    Published         : 2026-11-02T09:14:00Z
    Published SHA-256 : 9f2c1a...

    To install:  .\Clear-PrintSpoolerQueue.ps1 -Update -Force
```

Exit codes make this pollable:

| Code | Meaning |
|---:|---|
| 0 | Already current |
| 7 | A newer version exists — nothing was changed |
| 6 | GitHub could not be reached |

```powershell
.\Clear-PrintSpoolerQueue.ps1 -CheckUpdate
if ($LASTEXITCODE -eq 7) { .\Clear-PrintSpoolerQueue.ps1 -Update -Force }
```

---

## Three discovery strategies

Tried in order, so discovery keeps working when any one of them is unavailable:

| # | Source | Works when |
|---|---|---|
| 1 | `GET /repos/{owner}/{repo}/releases/latest` | The repository publishes releases. Authoritative: carries assets and checksums |
| 2 | `GET /repos/{owner}/{repo}/tags` | Only tags exist, no formal releases |
| 3 | `GET raw.githubusercontent.com/{owner}/{repo}/{branch}/Clear-PrintSpoolerQueue.ps1` | Neither exists — the version stamp is read from the branch source |

Strategy 3 parses the version with the **PowerShell AST** rather than a regex over the whole file, so
a version-looking string in a comment cannot fool it. This repository had no releases and no tags
before v4, and strategy 3 is what found `3.1.0` during development.

Each strategy's failure is recorded in `Errors[]` rather than being swallowed, so a rate-limited run
tells you *why* it fell through:

```
Discovery: releases: GitHub API rate limit or forbidden (403). Set PSQ_GITHUB_TOKEN, or rely on
the raw-file fallback.
```

### Channels

| Channel | Selects |
|---|---|
| `latest` *(default)* | The newest non-draft release |
| `stable` | The newest release not marked prerelease |
| `prerelease` | The newest release including prereleases |

`-IncludePrerelease` additionally makes `Test-NewerVersion` accept a prerelease as a valid target.
Without it, `4.1.0-rc1` is never considered newer than `4.0.0` — which is the right default for
production.

### Rate limits

Anonymous GitHub API access allows **60 requests per hour per source IP**. A fleet of 400 hosts
checking independently shares your NAT address and will hit that ceiling.

```powershell
$env:PSQ_GITHUB_TOKEN = 'ghp_...'      # 5000 req/h
```

Prefer the environment variable over `-GitHubToken`: a token on a command line is visible in process
listings and in scheduled-task definitions. The token is never written to the log, the report, or the
diagnostics bundle (the effective-config export redacts it).

---

## Applying an update

```powershell
.\Clear-PrintSpoolerQueue.ps1 -Update -Force
Update-SpoolerCleanup -Force                  # cmdlet form
Update-SpoolerCleanup -Force -ApplyNow        # install immediately rather than staging
```

Sequence:

1. Discover the latest version and its published SHA-256.
2. Download the assets to a temporary directory.
3. **Verify the SHA-256.** On mismatch the download is discarded, the installed copy is untouched,
   and the update fails with exit 9.
4. Parse the downloaded script. A file with syntax errors is discarded — a truncated or corrupt
   download must never be promoted.
5. Back up the current files to `<install>\backup-<version>-<stamp>\`.
6. Stage the new files in `<install>\.update\` with a `pending-update.json` manifest recording each
   file's size and SHA-256.
7. Apply, either now (`-ApplyNow`) or at the start of the next run.

---

## Staged self-update

A process cannot safely overwrite the file it is executing from. Windows locks the image, and even
where a rename succeeds, the running process is still executing the old bytes while the new file
sits on disk — so a half-applied update is indistinguishable from a successful one.

The tool therefore **stages** and swaps at start-up:

```
<install>\
├── Clear-PrintSpoolerQueue.ps1        <- currently executing
├── Clear-PrintSpoolerQueue.bat
└── .update\
    ├── Clear-PrintSpoolerQueue.ps1    <- verified, waiting
    ├── Clear-PrintSpoolerQueue.bat
    └── pending-update.json            <- manifest with per-file SHA-256
```

Both entry points apply the swap **before** running anything else:

- the `.ps1` calls `Invoke-PendingUpdateInternal`, which re-verifies each staged file's SHA-256
  against the manifest before moving it
- the `.bat` performs a plain `move`, which is safe because those bytes were verified twice already
  (at download against `checksums.txt`, and when written into `pending-update.json`)

Files are **moved**, not copied, so a half-applied update cannot be applied twice. If a file is
locked, the update stays staged and is applied on the next start — the message says so rather than
pretending to have succeeded.

> The `.bat` never replaces *itself*. `cmd.exe` re-reads a batch file by byte offset while it
> executes, so overwriting it mid-run produces garbage. The launcher's own update lands on the
> following start, and it tells you that.

### Rollback

The previous version is always kept:

```
<install>\backup-4.0.0-20261010-031500\
```

Rolling back is a copy:

```powershell
Copy-Item <install>\backup-4.0.0-20261010-031500\* <install>\ -Force
```

---

## Integrity verification

The default is to **refuse unverified bytes**:

| Situation | Behaviour |
|---|---|
| Release publishes `checksums.txt` | SHA-256 verified; mismatch discards the download |
| No checksums published | Update **refused** with an explanatory error |
| `-SkipChecksumVerification` | Installs anyway, records `VerifyResult = SkippedByOperator`, logs a warning |

`tools/Build.ps1` generates `checksums.txt` for every artifact, so a properly published release
always has one. If you are mirroring the tool internally, mirror `checksums.txt` with it.

`VerifyResult` appears in the update output and in the report:

| Value | Meaning |
|---|---|
| `ChecksumMatch` | Verified against the published hash |
| `ChecksumMismatch` | Rejected (the update failed) |
| `NoChecksumPublished` | Rejected unless overridden |
| `SkippedByOperator` | Installed without verification — you asked for it |

---

## Automatic updates

```powershell
.\Clear-PrintSpoolerQueue.ps1 -AutoUpdate -Force
```

`-AutoUpdate` performs a quiet check before the cleanup. It is designed for scheduled tasks:

- **Cooldown-limited** — at most one check per `Update.CooldownHours` (default 24), tracked in
  `update-state.json`
- **Never blocks** — a slow or failed check costs nothing and cannot delay the cleanup
- **Never fails the run** — an update error is logged at `DEBUG` and the cleanup proceeds
- **Stages rather than swaps** — so the running process is never disturbed

The scheduled task created by `Install-SpoolerCleanupTask` includes `-AutoUpdate` by default
(disable with `-NoAutoUpdate`).

### Air-gapped environments

Turn the whole mechanism off so it stops appearing in your logs:

```powershell
$env:PSQ_NO_UPDATE_CHECK = 'true'        # machine-wide via GPO preferences
```

or `Update.AutoCheck = $false` / `Update.Enabled = $false` in the configuration file.

---

## Versioning scheme

Semantic versioning: `MAJOR.MINOR.PATCH`, optionally `-prerelease` and `+build`.

Comparison is numeric per component, then "no prerelease beats prerelease", then ordinal on the
prerelease label. `4.10.0 > 4.9.0` — a lexical comparison would get that wrong, and it is asserted
by the test suite along with thirteen other pairs.

| Where | How |
|---|---|
| Installed version | `.\Clear-PrintSpoolerQueue.ps1 -ShowVersion` or `Get-SpoolerCleanupVersion` |
| Build stamp | `Get-SpoolerCleanupVersion \| Select-Object Version, Build` |
| Source of truth | `$Script:ScriptVersion` in `src/Shared/00-Constants.ps1` |

`tools/Build.ps1` stamps the version into the `.ps1`, the `.psd1`, the `.bat` and the release notes
from that single constant, and the build **fails** if any `@@TOKEN@@` survives — which is how v3's
README ("v3.0 Professional") drifted from its own scripts (3.1.0).

---

## Publishing a release

For maintainers:

```powershell
./tools/Build.ps1 -Version 4.1.0 -SigningThumbprint <thumbprint>
```

Produces in `dist/`:

| Artifact | Purpose |
|---|---|
| `Clear-PrintSpoolerQueue.ps1` | The standalone script |
| `Clear-PrintSpoolerQueue.bat` | The launcher, version-stamped |
| `PrintSpoolerCleanup-4.1.0.zip` | Module + script + launcher + config template + docs |
| `PrintSpoolerCleanup/` | The importable module |
| `checksums.txt` | SHA-256 for every artifact — **the updater verifies against this** |
| `RELEASE-NOTES.md` | Generated release body from `CHANGELOG.md` |

Upload everything including `checksums.txt`. An update published without it will be refused by
every client that tries to install it.

---

## Related pages

- [Security and Privacy](Security-and-Privacy)
- [For Developers](For-Developers) — the build pipeline
- [Scheduling and Automation](Scheduling-and-Automation)
- [Exit Codes](Exit-Codes)
