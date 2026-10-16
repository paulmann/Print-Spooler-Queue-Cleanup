// Tray implementation.
#include "tray.h"
#include "i18n.h"
#include "settings.h"
#include "spooler.h"
#include "util.h"
#include "main_window.h"
#include "about.h"
#include "theme.h"

namespace psqc { const wchar_t* Str(int id); }

enum TrayIds : int {
    IDM_TRAY_OPEN = 901, IDM_TRAY_JOBS, IDM_TRAY_QUEUE, IDM_TRAY_PRINTERS, IDM_TRAY_CLEAR,
    IDM_TRAY_PAUSE, IDM_TRAY_RESUME, IDM_TRAY_RESTART,
    IDM_TRAY_THEME_SYS, IDM_TRAY_THEME_LIGHT, IDM_TRAY_THEME_DARK,
    IDM_TRAY_AUTOSTART, IDM_TRAY_UPDATE, IDM_TRAY_ABOUT, IDM_TRAY_EXIT,
};

static NOTIFYICONDATAW g_nid{};
static bool g_present = false;
static int  g_jobs = 0;

bool TrayPresent() { return g_present; }

static void SetTip() {
    std::wstring tip = util::Fmt(L"%s — %s", psqc::Str(STR_APP_TITLE),
                                 util::Fmt(psqc::Str(STR_TRAY_JOBS), util::FormatInt(g_jobs).c_str()).c_str());
    wcsncpy_s(g_nid.szTip, tip.c_str(), _TRUNCATE);
}

void TrayCreate() {
    if (g_present || !g.hwndMain) return;
    g_nid = {};
    g_nid.cbSize = sizeof(g_nid);
    g_nid.hWnd = g.hwndMain;
    g_nid.uID = 1001;
    g_nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP | NIF_SHOWTIP;
    g_nid.uCallbackMessage = WM_PSQC_TRAYICON;
    g_nid.hIcon = util::LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
    SetTip();
    wcsncpy_s(g_nid.szInfoTitle, psqc::Str(STR_APP_TITLE), _TRUNCATE);
    g_present = Shell_NotifyIconW(NIM_ADD, &g_nid) != 0;
    if (g_present) {
        g_nid.uVersion = NOTIFYICON_VERSION_4;
        Shell_NotifyIconW(NIM_SETVERSION, &g_nid);
    }
}

void TrayDestroy() {
    if (!g_present) return;
    Shell_NotifyIconW(NIM_DELETE, &g_nid);
    g_present = false;
}

void TrayReskin() {
    if (!g_present) return;
    SetTip();
    g_nid.uFlags = NIF_TIP | NIF_ICON;
    g_nid.hIcon = util::LoadAppIcon(GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON));
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

void TrayUpdateCounts(int jobs) {
    g_jobs = jobs;
    if (g_present) {
        g_nid.uFlags = NIF_TIP;
        SetTip();
        Shell_NotifyIconW(NIM_MODIFY, &g_nid);
    }
}

void TrayNotify(int kind, const std::wstring& text) {
    if (!S().notifications || !g_present) return;
    g_nid.uFlags = NIF_INFO;
    wcsncpy_s(g_nid.szInfo, text.c_str(), _TRUNCATE);
    wcsncpy_s(g_nid.szInfoTitle, psqc::Str(STR_APP_TITLE), _TRUNCATE);
    g_nid.dwInfoFlags = kind == BK_OK ? NIIF_INFO : kind == BK_WARN ? NIIF_WARNING :
                        kind == BK_ERR ? NIIF_ERROR : NIIF_INFO;
    Shell_NotifyIconW(NIM_MODIFY, &g_nid);
}

static void BuildMenu(HMENU menu) {
    std::vector<spooler::PrinterJobs> snap;
    spooler::EnumerateAll(snap);
    bool anyPaused = spooler::AnyPaused(snap);

    AppendMenuW(menu, MF_STRING, IDM_TRAY_OPEN, psqc::Str(STR_TRAY_OPEN));
    SetMenuDefaultItem(menu, IDM_TRAY_OPEN, FALSE);
    AppendMenuW(menu, MF_STRING | MF_DISABLED, IDM_TRAY_JOBS,
                util::Fmt(psqc::Str(STR_TRAY_JOBS), util::FormatInt(g_jobs).c_str()).c_str());
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, IDM_TRAY_QUEUE, psqc::Str(STR_TRAY_QUEUE));
    AppendMenuW(menu, MF_STRING, IDM_TRAY_PRINTERS, psqc::Str(STR_TRAY_PRINTERS));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, IDM_TRAY_CLEAR, psqc::Str(STR_TRAY_CLEAR));
    AppendMenuW(menu, MF_STRING, anyPaused ? IDM_TRAY_RESUME : IDM_TRAY_PAUSE,
                psqc::Str(anyPaused ? STR_TRAY_RESUME : STR_TRAY_PAUSE));
    AppendMenuW(menu, MF_STRING, IDM_TRAY_RESTART, psqc::Str(STR_TRAY_RESTART));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    HMENU theme = CreatePopupMenu();
    AppendMenuW(theme, MF_STRING | (S().themeMode == THEME_SYSTEM ? MF_CHECKED : 0), IDM_TRAY_THEME_SYS, psqc::Str(STR_THEME_SYSTEM));
    AppendMenuW(theme, MF_STRING | (S().themeMode == THEME_LIGHT ? MF_CHECKED : 0), IDM_TRAY_THEME_LIGHT, psqc::Str(STR_THEME_LIGHT));
    AppendMenuW(theme, MF_STRING | (S().themeMode == THEME_DARK ? MF_CHECKED : 0), IDM_TRAY_THEME_DARK, psqc::Str(STR_THEME_DARK));
    AppendMenuW(menu, MF_POPUP, (UINT_PTR)theme, psqc::Str(STR_TRAY_THEME));
    AppendMenuW(menu, MF_STRING | (AutostartEnabled() ? MF_CHECKED : 0), IDM_TRAY_AUTOSTART, psqc::Str(STR_TRAY_AUTOSTART));
    AppendMenuW(menu, MF_STRING, IDM_TRAY_UPDATE, psqc::Str(STR_TRAY_UPDATE));
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING, IDM_TRAY_ABOUT, psqc::Str(STR_TRAY_ABOUT));
    AppendMenuW(menu, MF_STRING, IDM_TRAY_EXIT, psqc::Str(STR_TRAY_EXIT));
}

static void OnCommand(int id) {
    switch (id) {
        case IDM_TRAY_OPEN: MainWindowShowAndActivate(); break;
        case IDM_TRAY_QUEUE: {
            auto def = util::DefaultPrinterName();
            if (def.empty()) {
                std::vector<spooler::PrinterJobs> snap;
                spooler::EnumerateAll(snap);
                for (auto& p : snap) if (!p.jobs.empty()) { def = p.name; break; }
            }
            util::OpenPrinterQueueUI(def);
            break;
        }
        case IDM_TRAY_PRINTERS: util::OpenPrintersFolder(); break;
        case IDM_TRAY_CLEAR:
            MainWindowShowAndActivate();
            MainWindowBeginPurge(true);
            break;
        case IDM_TRAY_PAUSE: spooler::PauseAll(true); MainWindowRefreshNow(); break;
        case IDM_TRAY_RESUME: spooler::PauseAll(false); MainWindowRefreshNow(); break;
        case IDM_TRAY_RESTART: {
            std::wstring err;
            if (spooler::RestartService(err)) TrayNotify(BK_OK, psqc::Str(STR_MSG_RESTART_OK));
            else TrayNotify(BK_ERR, util::Fmt(psqc::Str(STR_MSG_STOP_FAIL), err.c_str()));
            MainWindowRefreshNow();
            break;
        }
        case IDM_TRAY_THEME_SYS: case IDM_TRAY_THEME_LIGHT: case IDM_TRAY_THEME_DARK: {
            int mode = id == IDM_TRAY_THEME_SYS ? THEME_SYSTEM : id == IDM_TRAY_THEME_LIGHT ? THEME_LIGHT : THEME_DARK;
            S().themeMode = mode;
            SettingsSave();
            ThemeSetMode(mode);
            MainWindowReskin();
            TrayReskin();
            break;
        }
        case IDM_TRAY_AUTOSTART: {
            bool on = !AutostartEnabled();
            std::wstring err;
            if (!AutostartSet(on, err)) TrayNotify(BK_ERR, err);
            break;
        }
        case IDM_TRAY_UPDATE: UpdateCheckAsync(true); break;
        case IDM_TRAY_ABOUT: AboutShow(); break;
        case IDM_TRAY_EXIT:
            TrayDestroy();
            DestroyWindow(g.hwndMain);
            break;
    }
}

void TrayHandleMessage(WPARAM wp, LPARAM lp);   // called from main window proc bridge
void TrayHandleMessage(WPARAM, LPARAM lp) {
    switch (LOWORD(lp)) {
        case WM_LBUTTONDBLCLK:
            MainWindowShowAndActivate();
            break;
        case WM_RBUTTONUP:
        case WM_CONTEXTMENU: {
            HMENU menu = CreatePopupMenu();
            BuildMenu(menu);
            SetForegroundWindow(g.hwndMain);
            POINT pt{};
            GetCursorPos(&pt);
            int cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, pt.x, pt.y, 0, g.hwndMain, nullptr);
            PostMessageW(g.hwndMain, WM_NULL, 0, 0);
            DestroyMenu(menu);
            if (cmd) OnCommand(cmd);
            break;
        }
        case NIN_BALLOONUSERCLICK:
            break;
    }
}
