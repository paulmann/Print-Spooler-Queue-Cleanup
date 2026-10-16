# Microsoft Store publication package (MSIX, full-trust Win32)

The app is a classic Win32 executable; the Store accepts it as a **full-trust**
(`runFullTrust`) MSIX package. Everything needed lives in `packaging/`:

```
packaging/
├─ AppxManifest.xml          template (identity placeholders marked with <<…>>)
├─ assets/
│  ├─ Square44x44Logo.png    (44×44)   ├─ StoreLogo.png        (50×50)
│  ├─ SmallTile71x71Logo.png (71×71)   ├─ Wide310x150Logo.png  (310×150)
│  ├─ Square150x150Logo.png  (150×150) ├─ LargeTile310x310Logo.png
│  └─ SplashScreen620x300.png (620×300)
└─ store-listing.md          listing text (EN + RU), keywords, categories
```

Assets are generated from `assets/icon-1024.png` by `tools/make_assets.py` — re-run it
whenever the icon changes.

## 1. Prerequisites

1. A **Microsoft Partner Center** developer account (one-time fee).
2. Reserve the app name (e.g. *Print Spooler Queue Cleanup*) → you receive
   `Publisher="CN=<<YOUR-PUBLISHER-SID>>"` and the package identity
   (`Name`, `PublisherDisplayName`) — put them into `AppxManifest.xml`.
3. Windows 10/11 SDK tools on the build machine: `makeappx.exe`, `signtool.exe`
   (from the Windows SDK; **not** needed on target machines).

## 2. Prepare the package layout

```
pkg/
├─ AppxManifest.xml
├─ SpoolerCleanup.exe        (release build, MSVC /MT, signed — see docs/SMARTSCREEN.md)
├─ assets/…                  (all PNGs above)
└─ LICENSE.txt               (PolyForm Noncommercial 1.0.0)
```

Notes:

* Keep `runFullTrust` + `EntryPoint="Windows.FullTrustApplication"` (already in the
  template) — this is what lets a classic exe run from a Store package.
* The app writes settings to `%APPDATA%` — allowed: Store full-trust apps get a normal
  per-user profile plus a virtualized local cache; our paths are user-writable.
* Autostart: full-trust apps may use the `Run` key (we do) or the Startup Task extension;
  the wizard's autostart toggle covers the UX, no extra extension required.
* The built-in self-updater should be **disabled in the Store build** (Store updates the
  package): build with `-DPSQC_STORE_BUILD` and guard `update::CheckForUpdate` calls —
  the guard macro is honored in `main.cpp`/`main_window.cpp` (see `PSQC_STORE_BUILD`
  notes in BUILD.md). *If you prefer keeping the in-app updater, declare
  `rescap:allowExternalContent` — not recommended.*

## 3. Pack & sign (local test)

```cmd
makeappx pack /d pkg /p SpoolerCleanup_5.0.0.0_x64.msix /l
signtool sign /fd SHA256 /f TestCert.pfx /p <pwd> SpoolerCleanup_5.0.0.0_x64.msix
```

For Partner Center you do **not** sign yourself: upload the `.msix` (or a `.zip` with the
package layout) and Microsoft signs it during certification.

## 4. Submit

1. Partner Center ▸ Your app ▸ *Packages* ▸ upload the MSIX.
2. Listing: use `packaging/store-listing.md` (description, features, keywords, category
   *Productivity / Utilities*), screenshots 1–8 (take on Windows 11, light + dark),
   store logo 300×300 (from `assets/icon-1024.png`).
3. Policies: the app is free; license = PolyForm Noncommercial (state it in the listing);
   no personal data collected (declare: no data collection, no telemetry — true);
   elevation: the app requests UAC only for explicit admin operations — declare
   "requires admin for some features" in notes for certification.
4. Certification notes field: mention full-trust Win32, printer-management capability
   rationale (system maintenance utility), and link to this repository + LICENSE.

## 5. Post-certification

* Store delivers updates: bump `Version` in `AppxManifest.xml` per release.
* Keep the GitHub release pipeline for non-Store users (signed exe).
* The in-app *Check for updates* in the Store build may link to the Store page instead of
  downloading (recommended UX).
