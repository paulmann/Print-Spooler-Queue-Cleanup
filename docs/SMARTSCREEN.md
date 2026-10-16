# SmartScreen / "Windows protected your PC" — why it happens and how to fix it

> Goal from the project owner: *«сделать так, чтобы при старте Windows не ругался на .exe»*.
> This document explains every layer that can "complain" about a new executable and the
> exact, legitimate ways to silence each layer.

## The layers that warn

| Layer | What it checks | Trigger |
|---|---|---|
| **SmartScreen reputation (Mark-of-the-Web)** | File reputation: signature + publisher + download volume | Any unsigned downloaded exe → blue "Windows protected your PC" |
| **SmartScreen / Defender AV scan** | Malware signatures & cloud heuristics | Actual malware or rare false positives |
| **UAC** | Manifest `requestedExecutionLevel` | Only when elevation is requested (our elevate-on-demand children) |
| **Zone identifier ADS** | `Zone.Identifier` alternate stream added by browsers | Downloads from Internet |

Our binary is clean by construction (no packer, no obfuscation, no injection, static link,
proper manifest, version info with company/product/copyright) — that minimizes AV
heuristics. What remains is **reputation**, and reputation is earned by a signature and/or
distribution volume.

## Fix levels (pick what applies to you)

### Level 0 — nothing (development / private use)
Right-click the file ▸ *Properties* ▸ check **Unblock** (removes Mark-of-the-Web), or in
the SmartScreen dialog click *More info* ▸ *Run anyway*. Fine for your own machines.

### Level 1 — Microsoft Store distribution (recommended, already prepared)
Apps installed from the **Microsoft Store** never trigger SmartScreen: the package is
signed by Microsoft and has no Mark-of-the-Web. Everything needed is in
[`docs/MICROSOFT-STORE.md`](./MICROSOFT-STORE.md) (`packaging/` holds the manifest template
and assets). This is the only channel where the warning disappears with **zero**
certificate cost.

### Level 2 — code-sign the binary (self-distribution)
1. Buy a certificate: **OV code-signing** (organization validated) or, better,
   **EV code-signing** (hardware token / cloud HSM).
2. Sign: `signtool sign /fd SHA256 /tr http://timestamp.acme.com /td SHA256 /sha1 <thumbprint> SpoolerCleanup.exe`
   (EV: `signtool sign /tr … /td SHA256 /f token.pfx …` per vendor docs).
3. **EV certificate** → SmartScreen reputation is granted *immediately* for every build
   signed with it. This is the definitive fix for GitHub-release distribution.
4. **OV certificate** → reputation accumulates per file over download volume; expect
   warnings for the first days/weeks.
5. Cheap modern alternative: **Microsoft Azure Trusted Signing** (paid per-month,
   identity-validated) — signed files get SmartScreen reputation via the Trusted Signing
   root; works with `signtool` + the Trusted Signing endpoint or `AzureSignTool`.
6. Re-sign **every** release; also sign the MSIX if you sideload it.

### Level 3 — reputation hygiene (free, slow)
Unsigned files accumulate reputation as more people run them and report "no problem"
(Defender cloud). Not a plan, just gravity. Do not rely on it.

## What we already do in-repo to keep heuristics quiet

* Authenticode-friendly PE: version resource with `CompanyName`, `FileDescription`,
  `LegalCopyright`, `OriginalFilename`, `ProductVersion` (missing version info is a
  classic SmartScreen/AV heuristic trigger).
* Application manifest with supported-OS GUIDs, `asInvoker`, DPI awareness — no
  "suspicious" requestedExecutionLevel jumps; elevation only via explicit child processes
  with user consent.
* No packers/protectors, no self-modifying code, no process injection; the self-update
  replace is a detached child doing `MoveFileEx` on files — the same pattern used by
  countless updaters, kept simple and readable.
* Network I/O limited to `https://api.github.com` + `https://github.com/…/releases/download/…`
  over WinHTTP with default TLS — no custom C2-looking traffic.
* Static linking: no DLL sideloading surface.

## Verification checklist before publishing a release

```cmd
signtool verify /pa /v SpoolerCleanup.exe        :: signature chain OK
signtool verify /pa /v <file> | findstr /i "sha256"
certutil -dump <file> | findstr /i "version"     :: version info present
```

And on a clean VM: download the signed release with Edge → run → expect **no** blue screen
when EV/Trusted Signing is used; with OV expect it only until reputation builds.
