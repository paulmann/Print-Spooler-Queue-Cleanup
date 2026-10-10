# Refactor & Productization Plan — Print Spooler Queue Cleanup v4.0.0

> Internal engineering plan. Percentages are completion against this plan, reported at each
> agent-swarm milestone. Baseline = upstream `paulmann/Print-Spooler-Queue-Cleanup` @ `a70a430` (v3.1.0).

## 0. Baseline audit (findings) — **DONE 100%**

| # | Finding (v3.1.0) | Severity | Fixed in v4 |
|---|---|---|---|
| A1 | `Spooler` stopped without stopping dependents (`Fax`, `PrintNotify`, `PrintWorkflowUserSvc*`) → stop can hang/fail on Server & Win11 | High | Yes |
| A2 | No selective cleanup (all `*.SHD`/`*.SPL` always deleted; cannot target a printer, job id, or age) | High | Yes |
| A3 | No backup before destructive delete; no restore | High | Yes |
| A4 | Locked-file handling: single attempt, no retry/backoff, no "file resurrected by spooler" detection | High | Yes |
| A5 | No diagnostics: cannot answer *why* the queue is stuck (job owner, age, size, driver, port, event log) | High | Yes |
| A6 | No repair path for the real-world root causes (corrupt job, bad driver, `PrintNotify` hang, Forms\cache, bad `DefaultSpoolDirectory`) | High | Yes |
| A7 | No scheduled/automated mode, no watchdog, no Task Scheduler installer | High | Yes |
| A8 | No update mechanism; users run stale scripts forever | Med | Yes |
| A9 | Help is thin: no comprehensive reference, no exit-code/env/config documentation in one place | Med | Yes |
| A10 | `$script:Stats.JobsCleared += [Math]::Max(SHD,SPL)` is a *guess*, not a real job count | Med | Yes |
| A11 | Sequential remote execution only — 200 workstations = 200 × (stop+3 s settle+start) | Med | Yes |
| A12 | No `-Credential`; remote runs as caller only | Med | Yes |
| A13 | No machine-readable output (JSON/CSV) → cannot be consumed by RMM/monitoring | Med | Yes |
| A14 | Logging: single file next to the script, no rotation, no retention, no size cap, writes into `Program Files` when installed there (fails) | Med | Yes |
| A15 | README claims "License: Public Domain" but `LICENSE` is **GPL-3.0** → legal/branding inconsistency | Med | Yes |
| A16 | README title/version drift ("v3.0 Professional" while scripts are 3.1.0); documented params don't match code | Med | Yes |
| A17 | No module form → cannot `Import-Module`, no discoverable cmdlets, no `FunctionsToExport` | Med | Yes |
| A18 | No tests, no CI, no build pipeline, no release checksums, no signing guidance | Med | Yes |
| A19 | No config file / env-var configuration → every site edit = code edit | Low | Yes |
| A20 | Batch: `pause` in scheduled context, `net stop`/`net start` locale quirks, no `/Elevate`, no update/help parity with PS | Low | Yes |
| A21 | `$global:LASTEXITCODE` assignment is a hack; dot-sourcing still leaks state; no `Throw`/exit contract | Low | Yes |
| A22 | No self-elevation, no non-interactive guard (hangs on `set /p` under Task Scheduler) | Low | Yes |
| A23 | No event-log integration actually implemented (README advertised it) | Low | Yes |

## 1. Architecture & contracts — **DONE 100%**

* Source-of-truth layout: `src/Shared`, `src/Private`, `src/Public` → assembled by `tools/Build.ps1`
  into **two** distributables that never drift:
  * `Clear-PrintSpoolerQueue.ps1` — single self-contained file (field technicians: copy one file).
  * `PrintSpoolerCleanup/PrintSpoolerCleanup.psd1|psm1` — importable module (DevOps/RMM).
* Contract: exit codes, log schema, JSON report schema, config schema — frozen in `docs/Exit-Codes.md`
  and enforced by tests.
* Token-based template injection (`@@TOKEN@@`) for the remote payload — no `-f` brace escaping bugs.

## 2. Milestones

| ID | Milestone | Weight | Status |
|----|-----------|-------:|--------|
| M0 | Recon + baseline audit | 5% | ✅ 100% |
| M1 | Architecture, plan, contracts | 5% | ✅ 100% |
| M2 | Core engine (constants, logging, config, service orchestration, selective cleanup, retry, backup, restore) | 22% | ✅ 100% |
| M3 | Diagnostics + Repair + Health cmdlets | 8% | ✅ 100% |
| M4 | Update subsystem (GitHub Releases, checksum, staged self-update, cooldown) | 7% | ✅ 100% |
| M5 | CLI/UX: comprehensive help, progress, colors, machine output, exports, elevation | 8% | ✅ 100% |
| M6 | Batch launcher v4 (PS7-aware, elevate, pending-update swap, full parity) | 5% | ✅ 100% |
| M7 | README + CHANGELOG + docs set | 10% | ✅ 100% |
| M8 | Wiki (English, GitHub-ready; planned 18 pages, delivered 23 content pages + Home + `_Sidebar`) | 15% | ✅ 100% |
| M9 | Production infra: Build.ps1, Deploy.ps1, CI workflows, templates, .gitignore, manifest | 7% | ✅ 100% |
| M10 | Tests (parser gate + unit + contract) executed on PowerShell 7.6.6 | 6% | ✅ 100% |
| M11 | Final validation, packaging, handover report (re-verified on a clean checkout, Linux + PowerShell 7.6.6, 2026-10-10) | 2% | ✅ 100% |

**Plan completion: 100% (all 12 milestones).** Final gate: 9/9 stages, **715 checks**, green on both
required paths (direct `tests/Run-Tests.ps1` and nested via `tools/Build.ps1`) on a clean checkout —
Linux / PowerShell 7.6.6, 2026-10-10. The completion sweep after the first "100%" found defects
21–48 (see `HANDOVER.md` §4); each produced a permanent gate check, and the total grew from 697 to
715 accordingly.

## 3. New feature set (what was missing)

**Cleanup precision** — `-PrinterName` / `-ExcludePrinterName` / `-JobId` / `-DocumentName` /
`-OlderThan` / `-NewerThan` / `-LargerThanBytes`, WMI `Win32_PrintJob` inventory with `.SHD`-header fallback.

**Safety** — `-Backup` + `Restore-SpoolerBackup` (zip + JSON manifest + SHA-256), `-ArchivePath`,
retry with backoff, post-restart "resurrection" verification, `-RequireConfirmation`,
non-interactive guard, spool-dir sanity guards (never recursive, never outside allow-list).

**Operations** — `-Diagnostic`, `-Repair`, `Test-PrintSpoolerHealth`, `-Watchdog`, `-SelfTest`,
dependent-service handling, `PrintNotify` recovery, Forms cache purge, optional driver-store purge,
spool-directory relocation.

**Scale** — `-ThrottleLimit` (PS7 `ForEach-Object -Parallel`, PS5.1 runspace fallback),
`-CimSession` reuse, `-Credential`, WinRM pre-flight probe, per-computer isolation (one failure never aborts the fleet).

**Observability** — leveled rotating log, JSON log mode, `Export-*` (JSON/CSV/HTML/Text),
Windows Event Log source `PrintSpoolerCleanup`, `-MachineReadableJson` for RMM.

**Lifecycle** — `Get-SpoolerCleanupUpdate`, `Update-SpoolerCleanup`, `-CheckUpdate`, `-Update`,
`-AutoUpdate` (24 h cooldown + state file + `--updated` loop guard), `Install-SpoolerCleanupTask`,
`Uninstall-SpoolerCleanupTask`, `Get-SpoolerCleanupVersion`.

**Ergonomics** — `Get-SpoolerCleanupHelp` (comprehensive, sectioned, searchable), `-Ascii` mode,
`-NoColor`, config file + environment variables, `-ConfigPath`, `-SaveConfig`, `-Elevate`,
`-VerifySignature`, `-PassThru`, `-Quiet`, `-NoProfile`.

## 4. Definition of Done

1. `tools/Build.ps1` produces both distributables with zero parser errors on PS 5.1 **and** 7.6.
2. `tests/Run-Tests.ps1` passes 100% (parser gate, unit, contract, batch lint).
3. Every documented switch exists in code; every code switch is documented (test-enforced).
4. README + wiki in English (planned 18 pages, delivered 23 content pages), no version drift,
   license consistent with `LICENSE`.
5. `dist/` contains ZIP + `checksums.txt` + release notes, ready to publish on GitHub Releases.
