#pragma once
// ---------------------------------------------------------------------------
// Print Spooler Queue Cleanup v5 — central application header.
// Pure Win32 / C++ (no .NET, no 3rd-party libs, no runtime deps).
// ---------------------------------------------------------------------------
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif
#ifndef NOMINMAX
#define NOMINMAX
#endif
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0A00   // Windows 10+
#endif
#ifndef NTDDI_VERSION
#define NTDDI_VERSION NTDDI_WIN10
#endif

#include <windows.h>
#include <windowsx.h>
#include <commctrl.h>
#include <shellapi.h>
#include <shlobj.h>
#include <shobjidl.h>
#include <winspool.h>
#include <winhttp.h>
#include <dwmapi.h>
#include <psapi.h>

#include <cstdarg>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <cwchar>
#include <string>
#include <string_view>
#include <vector>
#include <map>
#include <memory>
#include <functional>
#include <atomic>
#include <mutex>
#include <thread>
#include <optional>
#include <algorithm>

#include "version.h"

#ifdef _MSC_VER
#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "winspool.lib")
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "advapi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "user32.lib")
#endif

// ---- Custom window messages -----------------------------------------------
enum : UINT {
    WM_PSQC_TRAYICON      = WM_APP + 1,   // tray notifications (lParam = mouse msg)
    WM_PSQC_WORK_DONE     = WM_APP + 2,   // worker finished: wParam = task id, lParam = result code
    WM_PSQC_WORK_LOG      = WM_APP + 3,   // worker text: lParam = wchar_t* (caller frees)
    WM_PSQC_REFRESH       = WM_APP + 4,   // request list refresh
    WM_PSQC_THEME_CHANGED = WM_APP + 5,   // theme (re)applied
    WM_PSQC_LANG_CHANGED  = WM_APP + 6,   // language changed
    WM_PSQC_SINGLETON     = WM_APP + 7,   // second instance asks to show UI
    WM_PSQC_DL_PROGRESS   = WM_APP + 8,   // wParam = percent 0..100
    WM_PSQC_BALLOON       = WM_APP + 9,   // wParam = BalloonKind, lParam = wchar_t* payload
    WM_PSQC_TIMER_REFRESH = WM_APP + 10,
    WM_PSQC_UPDATE_CHECK_DONE = WM_APP + 11,
    WM_PSQC_UPDATE_DL_DONE    = WM_APP + 12,
};

enum BalloonKind : WPARAM {
    BK_INFO = 0, BK_OK = 1, BK_WARN = 2, BK_ERR = 3, BK_UPDATE = 4,
};

// Worker task ids posted with WM_PSQC_WORK_DONE
enum TaskId : WPARAM {
    TASK_PURGE_ALL = 1,
    TASK_DELETE_SEL = 2,
    TASK_RESTART_SPOOLER = 3,
    TASK_RESTORE = 4,
    TASK_UPDATE_CHECK = 5,
    TASK_UPDATE_DOWNLOAD = 6,
};

// ---- Global application state (defined in main.cpp) ------------------------
struct AppState {
    HINSTANCE hInst           = nullptr;
    HWND      hwndMain        = nullptr;
    HWND      hwndWizard      = nullptr;
    HWND      hwndSettings    = nullptr;
    HWND      hwndAbout       = nullptr;
    bool      elevated        = false;
    bool      portable        = false;
    bool      cliMode         = false;
    bool      quitRequested   = false;
    UINT      timerRefresh    = 0;
    int       liveJobCount    = 0;   // updated by refresh workers, shown in tray/status
};
extern AppState g;

// Forward decls (implemented across modules)
namespace psqc {
void ApplyThemeToWindow(HWND hwnd);          // theme.cpp
const wchar_t* Str(int id);                  // i18n.cpp
}
