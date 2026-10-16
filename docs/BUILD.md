# Building Print Spooler Queue Cleanup v5

The product is a **single static Win32 executable**. Two supported toolchains:

| Toolchain | Where | Notes |
|---|---|---|
| MSVC (VS 2022 / Build Tools) | Windows | `/std:c++latest` (C++26-ready), `/MT`, recommended for release |
| MinGW-w64 (GCC ≥ 12, posix threads) | Windows or Linux cross | used by CI and dev boxes |

No third-party libraries, no CMake superbuild, no package managers. The resource script
(`res/app.rc`) embeds the icon, the application manifest and the version block.

## MSVC (recommended for production)

From a *Developer Command Prompt* (x64):

```cmd
tools\build_msvc.bat
```

The script compiles every `src\*.cpp`, compiles `res\app.rc` with `rc.exe`, links with
`/MT /SUBSYSTEM:WINDOWS` and produces `dist\SpoolerCleanup.exe`.
Flags: `/std:c++latest /W4 /utf-8 /O2 /GS /guard:cf /DYNAMICBASE /NXCOMPAT /HIGHENTROPYVA`.

## MinGW-w64 (Linux cross / CI)

```bash
sudo apt install g++-mingw-w64-x86-64-posix   # POSIX thread model (std::thread)
bash tools/build_mingw.sh
# -> dist/SpoolerCleanup.exe  (~1.1 MB, fully static)
```

Override with `CXX=… WINDRES=… STD=c++2b` if needed. The build is `-static
-static-libgcc -static-libstdc++`: the result imports only Windows system DLLs
(verify: `x86_64-w64-mingw32-objdump -p dist/SpoolerCleanup.exe | grep "DLL Name"`).

## Unit tests (platform-independent logic)

```bash
bash tools/run_native_tests.sh     # host g++ or clang++; no Windows needed
```

Covers: integrity of all 30 translation tables (145 strings each, placeholder parity with
English, unique native names), semantic-version comparison, JSON mini-parser/escaper.

## CI

`.github/workflows/ci.yml` runs on every push/PR:

1. `windows-latest` MSVC build + artifact upload;
2. `ubuntu-latest` MinGW cross build + native tests + release asset packaging.

Release assets are named `SpoolerCleanup-<ver>-x64.exe` — the updater looks for exactly
this pattern (`*x64*.exe`) in the latest GitHub Release.

## Reproducibility checklist for release builds

* [ ] bump `PSQC_VERSION_*` in `src/version.h` (single source of truth: rc, About, MSIX, updater);
* [ ] `tools/make_assets.py` regenerated `res/app.ico` + `packaging/assets/*` if the icon changed;
* [ ] native tests green;
* [ ] MSVC build with `/MT` (no CRT redistribute);
* [ ] sign the binary (see `docs/SMARTSCREEN.md`) before publishing;
* [ ] attach to a GitHub Release with tag `v<major>.<minor>.<patch>`.
