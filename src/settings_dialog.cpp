// Settings dialog implementation.
#include "settings_dialog.h"
#include "controls.h"
#include "theme.h"
#include "i18n.h"
#include "settings.h"
#include "spooler.h"
#include "util.h"
#include "tray.h"
#include "wizard.h"
#include "about.h"
#include "main_window.h"

namespace psqc { const wchar_t* Str(int id); }

enum SdIds : int {
    IDS_SEC_APP = 401, IDS_SEC_BEH, IDS_SEC_LOG, IDS_SEC_DATA,
    IDS_T_TRAY = 410, IDS_T_MINTRAY, IDS_T_STARTMIN, IDS_T_NOTIFY, IDS_T_CONFIRM,
    IDS_T_AUTOUPD, IDS_T_ELEV, IDS_T_BACKUP, IDS_T_AUTOSTART,
    IDS_L_THEME = 430, IDS_L_LANG, IDS_L_LOG, IDS_L_REFRESH,
    IDS_C_THEME = 440, IDS_C_LANG, IDS_C_LOG, IDS_E_REFRESH,
    IDS_B_LOGS = 450, IDS_B_BACKUPS, IDS_B_RESTORE, IDS_B_UNINSTALL,
    IDS_NOTE = 455,
    IDS_T1 = 460, IDS_T2, IDS_T3, IDS_T4, IDS_T5, IDS_T6, IDS_T7, IDS_T8, IDS_T9,
    IDS_L_FONT = 470, IDS_L_SIZE,
    IDS_C_FONT = 472, IDS_C_SIZE,
};

struct SD {
    HWND hwnd = nullptr;
    HWND secApp, secBeh, secLog, secData;
    HWND lTheme, lLang, lLog, lRefresh, lFont, lSize;
    HWND cTheme, cLang, cLog, eRefresh, cFont, cSize;
    HWND tTray, tMinTray, tStartMin, tNotify, tConfirm, tAutoUpd, tElev, tBackup, tAutoStart;
    HWND bLogs, bBackups, bRestore, bUninstall;
    HWND note;
    HWND tl[9] = {};
};
static SD d;

// Список семейств, фактически попавших в комбобокс шрифта (только установленные).
static std::vector<std::wstring> g_sdFaces;
static bool FaceInstalled(const wchar_t* face) {
    HDC dc = GetDC(nullptr);
    HFONT f = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                          OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                          DEFAULT_PITCH | FF_SWISS, face);
    HGDIOBJ old = SelectObject(dc, f);
    wchar_t actual[LF_FACESIZE]{};
    GetTextFaceW(dc, LF_FACESIZE, actual);
    SelectObject(dc, old);
    DeleteObject(f);
    ReleaseDC(nullptr, dc);
    return _wcsicmp(actual, face) == 0;
}

static void FillCombos() {
    SendMessageW(d.cTheme, CB_RESETCONTENT, 0, 0);
    SendMessageW(d.cTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_SYSTEM));
    SendMessageW(d.cTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_LIGHT));
    SendMessageW(d.cTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_DARK));
    SendMessageW(d.cTheme, CB_SETCURSEL, S().themeMode, 0);
    SendMessageW(d.cLang, CB_RESETCONTENT, 0, 0);
    SendMessageW(d.cLang, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_SYSTEM));
    for (int i = 0; i < LangCount(); ++i) SendMessageW(d.cLang, CB_ADDSTRING, 0, (LPARAM)Langs()[i].native);
    int cur = 0;
    if (!S().lang.empty())
        for (int i = 0; i < LangCount(); ++i)
            if (util::IEquals(Langs()[i].tag, S().lang)) { cur = i + 1; break; }
    SendMessageW(d.cLang, CB_SETCURSEL, cur, 0);
    SendMessageW(d.cLog, CB_RESETCONTENT, 0, 0);
    SendMessageW(d.cLog, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_LOG_OFF));
    SendMessageW(d.cLog, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_LOG_ERRORS));
    SendMessageW(d.cLog, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_LOG_ALL));
    SendMessageW(d.cLog, CB_SETCURSEL, S().logLevel, 0);
    SendMessageW(d.cFont, CB_RESETCONTENT, 0, 0);
    g_sdFaces.clear();
    const wchar_t* cand[] = { L"Segoe UI", L"Segoe UI Variable Text", L"Inter", L"Arial",
                              L"Calibri", L"Tahoma", L"Verdana", L"Trebuchet MS" };
    std::wstring want = S().fontFace.empty() ? std::wstring(L"Segoe UI") : S().fontFace;
    int selFont = 0;
    for (const wchar_t* f : cand) {
        if (!FaceInstalled(f)) continue;
        int idx = (int)g_sdFaces.size();
        g_sdFaces.push_back(f);
        SendMessageW(d.cFont, CB_ADDSTRING, 0, (LPARAM)f);
        if (_wcsicmp(f, want.c_str()) == 0) selFont = idx;
    }
    SendMessageW(d.cFont, CB_SETCURSEL, selFont, 0);
    SendMessageW(d.cSize, CB_RESETCONTENT, 0, 0);
    const int sizes[] = { 8, 9, 10, 11, 12 };
    int selSize = 2;
    for (int i = 0; i < 5; ++i) {
        wchar_t b[8]{};
        swprintf(b, 8, L"%d", sizes[i]);   // _itow_s takes 4 args on MSVC
        SendMessageW(d.cSize, CB_ADDSTRING, 0, (LPARAM)b);
        if (sizes[i] == S().fontSize) selSize = i;
    }
    SendMessageW(d.cSize, CB_SETCURSEL, selSize, 0);
}

static void ApplyTexts() {
    SetWindowTextW(d.hwnd, psqc::Str(STR_DL_SETTINGS));
    SetWindowTextW(d.secApp, psqc::Str(STR_SET_APPEARANCE));
    SetWindowTextW(d.secBeh, psqc::Str(STR_SET_BEHAVIOR));
    SetWindowTextW(d.secLog, psqc::Str(STR_SET_LOGGING));
    SetWindowTextW(d.secData, psqc::Str(STR_SET_DATA));
    SetWindowTextW(d.lTheme, psqc::Str(STR_SET_THEME));
    SetWindowTextW(d.lLang, psqc::Str(STR_SET_LANGUAGE));
    SetWindowTextW(d.lLog, psqc::Str(STR_SET_LOGGING));
    SetWindowTextW(d.lRefresh, psqc::Str(STR_SET_REFRESH));
    SetWindowTextW(d.lFont, psqc::Str(STR_SET_FONT));
    SetWindowTextW(d.lSize, psqc::Str(STR_SET_FONTSIZE));
    CtrlSetStrId(d.bLogs, STR_BTN_OPEN_LOGS);
    CtrlSetStrId(d.bBackups, STR_BTN_OPEN_BACKUPS);
    CtrlSetStrId(d.bRestore, STR_BTN_RESTORE);
    CtrlSetStrId(d.bUninstall, STR_BTN_UNINSTALL);
    SetWindowTextW(d.note, psqc::Str(STR_SET_PORTABLE_NOTE));
    ShowWindow(d.note, g.portable ? SW_SHOW : SW_HIDE);
    if (d.tl[0]) {
        const int ids[9] = { STR_SET_TRAY, STR_SET_MIN_TRAY, STR_SET_START_MIN, STR_SET_NOTIFY, STR_SET_CONFIRM,
                             STR_SET_AUTOUPD, STR_SET_ELEVATE, STR_SET_BACKUP, STR_SET_AUTOSTART };
        for (int i = 0; i < 9; ++i) SetWindowTextW(d.tl[i], psqc::Str(ids[i]));
    }
    FillCombos();
    CtrlRelabelAll();
}

static void ApplyDlgFonts();   // определение ниже, рядом с SettingsProc

void SettingsReskin() {
    if (!d.hwnd) return;
    ThemeDarkTitlebar(d.hwnd, Th().dark);
    ThemeDarkControls(d.hwnd, Th().dark);
    ApplyTexts();
    ApplyDlgFonts();
    InvalidateRect(d.hwnd, nullptr, TRUE);
}

// ---- restore dialog -------------------------------------------------------------
struct RS { HWND hwnd = nullptr, list = nullptr, bRestore = nullptr, bCancel = nullptr; std::vector<spooler::BackupInfo> items; };
static RS r;

static void FillRestoreList() {
    r.items = spooler::ListBackups();
    SendMessageW(r.list, LB_RESETCONTENT, 0, 0);
    for (auto& b : r.items) {
        std::wstring line = util::Fmt(L"%s  —  %s, %s",
            util::FormatSystemTimeUser(b.when).c_str(),
            util::FormatInt(b.fileCount).c_str(),
            util::FormatBytes(b.totalSize).c_str());
        SendMessageW(r.list, LB_ADDSTRING, 0, (LPARAM)line.c_str());
    }
}

static LRESULT CALLBACK RestoreProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_ERASEBKGND: { RECT rc{}; GetClientRect(hwnd, &rc); FillRect((HDC)wp, &rc, Th().brBg); return 1; }
        case WM_CTLCOLORSTATIC: case WM_CTLCOLORLISTBOX:
            return (LRESULT)ThemeCtlColorBg((HDC)wp);
        case WM_DRAWITEM:
            if (CtrlHandleDrawItem((DRAWITEMSTRUCT*)lp)) return TRUE;
            return 0;
        case WM_COMMAND: {
            int id = LOWORD(wp);
            if (HIWORD(wp) == BN_CLICKED && id == 502) { DestroyWindow(hwnd); return 0; }
            if (HIWORD(wp) == BN_CLICKED && id == 501) {
                int sel = (int)SendMessageW(r.list, LB_GETCURSEL, 0, 0);
                if (sel < 0 || sel >= (int)r.items.size()) return 0;
                SetCursor(LoadCursorW(nullptr, IDC_WAIT));
                spooler::PurgeResult res;
                bool ok = spooler::RestoreSpoolFiles(r.items[sel].path, res);
                MessageBoxW(hwnd,
                    ok ? util::Fmt(psqc::Str(STR_MSG_RESTORE_OK), util::FormatInt(res.restoredFiles).c_str()).c_str()
                       : util::Fmt(psqc::Str(STR_MSG_RESTORE_FAIL), res.error.c_str()).c_str(),
                    psqc::Str(STR_RESTORE_TITLE), ok ? MB_ICONINFORMATION : MB_ICONERROR);
                if (ok && g.hwndMain) PostMessageW(g.hwndMain, WM_PSQC_REFRESH, 0, 0);
                return 0;
            }
            return 0;
        }
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_DESTROY: r.hwnd = nullptr; return 0;   // иначе диалог мёртв после первого закрытия
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

static void RestoreShow() {
    if (r.hwnd && IsWindow(r.hwnd)) { SetForegroundWindow(r.hwnd); return; }
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = RestoreProc;
    wc.hInstance = g.hInst;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.lpszClassName = L"SpoolerCleanupRestoreClass";
    RegisterClassExW(&wc);
    int dpi = util::DpiOf(g.hwndMain);
    auto Sc = [&](int v) { return MulDiv(v, dpi, 96); };
    r.hwnd = CreateWindowExW(0, L"SpoolerCleanupRestoreClass", psqc::Str(STR_RESTORE_TITLE),
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, Sc(460), Sc(330),
                             d.hwnd, nullptr, g.hInst, nullptr);
    ThemeDarkTitlebar(r.hwnd, Th().dark);
    util::SetClientSize(r.hwnd, Sc(460), Sc(330));   // Sc(460) x Sc(330) is the client area
    HWND hint = CreateWindowExW(0, L"STATIC", psqc::Str(STR_RESTORE_HINT), WS_CHILD | WS_VISIBLE, Sc(16), Sc(12), Sc(430), Sc(20), r.hwnd, nullptr, g.hInst, nullptr);
    r.list = CreateWindowExW(WS_EX_CLIENTEDGE, L"LISTBOX", L"", WS_CHILD | WS_VISIBLE | LBS_HASSTRINGS | WS_VSCROLL, Sc(16), Sc(36), Sc(428), Sc(210), r.hwnd, nullptr, g.hInst, nullptr);
    r.bRestore = CreatePsqcButton(r.hwnd, 501, STR_BTN_RESTORE_SEL, GLYPH_CHECK, BK_Accent);
    r.bCancel = CreatePsqcButton(r.hwnd, 502, STR_BTN_CLOSE, 0, BK_Standard);
    SendMessageW(hint, WM_SETFONT, (WPARAM)Th().font, TRUE);
    SendMessageW(r.list, WM_SETFONT, (WPARAM)Th().font, TRUE);
    ThemeDarkControls(r.hwnd, Th().dark);
    SetWindowPos(r.bRestore, nullptr, Sc(16), Sc(256), Sc(150), Sc(32), SWP_NOZORDER);
    SetWindowPos(r.bCancel, nullptr, Sc(354), Sc(256), Sc(90), Sc(32), SWP_NOZORDER);
    FillRestoreList();
    util::CenterWindow(r.hwnd, d.hwnd);
    ShowWindow(r.hwnd, SW_SHOW);
    UpdateWindow(r.hwnd);
#ifdef PSQC_QA_DELAY
    Sleep(700);
#endif
    util::QaShot(r.hwnd, L"restore");
}

// ---- settings dialog --------------------------------------------------------------
static void ApplyDlgFonts() {
    if (!d.hwnd) return;
    for (HWND h : { d.secApp, d.secBeh, d.secLog, d.secData })
        SendMessageW(h, WM_SETFONT, (WPARAM)Th().fontBold, TRUE);
    HWND fonted[] = { d.lTheme, d.lLang, d.lLog, d.lRefresh, d.lFont, d.lSize, d.note,
                      d.cTheme, d.cLang, d.cLog, d.cFont, d.cSize, d.eRefresh,
                      d.tTray, d.tMinTray, d.tStartMin, d.tNotify, d.tConfirm, d.tAutoUpd, d.tElev, d.tBackup, d.tAutoStart,
                      d.tl[0], d.tl[1], d.tl[2], d.tl[3], d.tl[4], d.tl[5], d.tl[6], d.tl[7], d.tl[8] };
    for (HWND h : fonted) SendMessageW(h, WM_SETFONT, (WPARAM)Th().font, TRUE);
}

static LRESULT CALLBACK SettingsProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_ERASEBKGND: { RECT rc{}; GetClientRect(hwnd, &rc); FillRect((HDC)wp, &rc, Th().brBg); return 1; }
        case WM_CTLCOLORSTATIC:
            return (LRESULT)ThemeCtlColorBg((HDC)wp);
        case WM_CTLCOLOREDIT:
            // «Автообновление списка»: в тёмной теме текст был чёрным на чёрном.
            SetBkColor((HDC)wp, Th().surface2);
            SetTextColor((HDC)wp, Th().text);
            return (LRESULT)Th().brSurface2;
        case WM_DRAWITEM:
            if (CtrlHandleDrawItem((DRAWITEMSTRUCT*)lp)) return TRUE;
            return 0;
        case WM_COMMAND: {
            int id = LOWORD(wp);
            int code = HIWORD(wp);
            if (code == BN_CLICKED) {
                auto Toggle = [&](HWND h, bool Settings::*field, void (*side)() = nullptr) {
                    bool on = !CtrlChecked(h);
                    S().*field = on;
                    CtrlSetChecked(h, on);
                    SettingsSave();
                    if (side) side();
                };
                switch (id) {
                    case IDS_T1: Toggle(d.tTray, &Settings::trayIcon, [] { if (S().trayIcon) TrayCreate(); else TrayDestroy(); }); return 0;
                    case IDS_T2: Toggle(d.tMinTray, &Settings::minimizeToTray); return 0;
                    case IDS_T3: Toggle(d.tStartMin, &Settings::startMinimized); return 0;
                    case IDS_T4: Toggle(d.tNotify, &Settings::notifications); return 0;
                    case IDS_T5: Toggle(d.tConfirm, &Settings::confirmClearAll); return 0;
                    case IDS_T6: Toggle(d.tAutoUpd, &Settings::autoUpdate); return 0;
                    case IDS_T7: Toggle(d.tElev, &Settings::autoElevate); return 0;
                    case IDS_T8: Toggle(d.tBackup, &Settings::backupBeforePurge); return 0;
                    case IDS_T9: {
                        bool on = !CtrlChecked(d.tAutoStart);
                        std::wstring err;
                        if (AutostartSet(on, err)) CtrlSetChecked(d.tAutoStart, on);
                        else { CtrlSetChecked(d.tAutoStart, !on); MessageBoxW(hwnd, err.c_str(), psqc::Str(STR_APP_TITLE), MB_ICONERROR); }
                        return 0;
                    }
                    case IDS_B_LOGS: util::OpenFolder(util::LogsDir(false)); return 0;
                    case IDS_B_BACKUPS: util::OpenFolder(util::BackupRootDir(false)); return 0;
                    case IDS_B_RESTORE: RestoreShow(); return 0;
                    case IDS_B_UNINSTALL: UninstallStart(false); return 0;
                }
            }
            if (code == CBN_SELCHANGE) {
                if ((HWND)lp == d.cTheme) {
                    int sel = (int)SendMessageW(d.cTheme, CB_GETCURSEL, 0, 0);
                    if (sel >= 0 && sel <= 2) {
                        S().themeMode = sel; SettingsSave();
                        ThemeSetMode(sel);
                        MainWindowReskin(); TrayReskin(); SettingsReskin();
                    }
                    return 0;
                }
                if ((HWND)lp == d.cLang) {
                    int sel = (int)SendMessageW(d.cLang, CB_GETCURSEL, 0, 0);
                    S().lang = sel <= 0 ? std::wstring() : Langs()[sel - 1].tag;
                    SettingsSave();
                    SetLanguage(S().lang.c_str());
                    MainWindowReskin(); TrayReskin(); SettingsReskin();
                    PostMessageW(g.hwndMain, WM_PSQC_LANG_CHANGED, 0, 0);
                    if (g.hwndAbout) PostMessageW(g.hwndAbout, WM_PSQC_LANG_CHANGED, 0, 0);
                    return 0;
                }
                if ((HWND)lp == d.cLog) {
                    int sel = (int)SendMessageW(d.cLog, CB_GETCURSEL, 0, 0);
                    if (sel >= 0 && sel <= 2) { S().logLevel = sel; SettingsSave(); util::SetLogLevel(sel); }
                    return 0;
                }
                if ((HWND)lp == d.cFont) {
                    int sel = (int)SendMessageW(d.cFont, CB_GETCURSEL, 0, 0);
                    if (sel >= 0 && sel < (int)g_sdFaces.size()) {
                        S().fontFace = g_sdFaces[sel];
                        SettingsSave();
                        ThemeSetMode(S().themeMode);   // пересборка шрифтов темы
                        MainWindowReskin(); TrayReskin(); SettingsReskin();
                        if (g.hwndAbout) AboutReskin();
                    }
                    return 0;
                }
                if ((HWND)lp == d.cSize) {
                    int sel = (int)SendMessageW(d.cSize, CB_GETCURSEL, 0, 0);
                    const int sizes[] = { 8, 9, 10, 11, 12 };
                    if (sel >= 0 && sel < 5) {
                        S().fontSize = sizes[sel];
                        SettingsSave();
                        ThemeSetMode(S().themeMode);
                        MainWindowReskin(); TrayReskin(); SettingsReskin();
                        if (g.hwndAbout) AboutReskin();
                    }
                    return 0;
                }
            }
            if (code == EN_CHANGE && (HWND)lp == d.eRefresh) {
                wchar_t buf[8]{};
                GetWindowTextW(d.eRefresh, buf, 8);
                int v = _wtoi(buf);
                if (v < 0) v = 0;
                if (v > 60) v = 60;
                if (v != S().refreshSec) {
                    S().refreshSec = v; SettingsSave();
                    if (g.timerRefresh) { KillTimer(g.hwndMain, 1); g.timerRefresh = 0; }
                    if (v > 0 && g.hwndMain) g.timerRefresh = SetTimer(g.hwndMain, 1, v * 1000, nullptr);
                }
                return 0;
            }
            return 0;
        }
        case WM_CLOSE: DestroyWindow(hwnd); return 0;
        case WM_DESTROY: g.hwndSettings = nullptr; return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void SettingsShow() {
    if (g.hwndSettings) { SetForegroundWindow(g.hwndSettings); return; }
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = SettingsProc;
    wc.hInstance = g.hInst;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = util::LoadAppIcon(32, 32);
    wc.lpszClassName = PSQC_DIALOG_CLASS;
    RegisterClassExW(&wc);
    int dpi = util::DpiOf(g.hwndMain);
    auto Sc = [&](int v) { return MulDiv(v, dpi, 96); };
    int W = Sc(520), H = Sc(710);
    d = SD{};
    d.hwnd = CreateWindowExW(0, PSQC_DIALOG_CLASS, psqc::Str(STR_DL_SETTINGS),
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, W, H,
                             g.hwndMain, nullptr, g.hInst, nullptr);
    g.hwndSettings = d.hwnd;
    util::SetClientSize(d.hwnd, W, H);   // W x H is the client area
    ThemeDarkTitlebar(d.hwnd, Th().dark);

    auto mkSec = [&](int id) { return CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)id, g.hInst, nullptr); };
    d.secApp = mkSec(IDS_SEC_APP); d.secBeh = mkSec(IDS_SEC_BEH); d.secLog = mkSec(IDS_SEC_LOG); d.secData = mkSec(IDS_SEC_DATA);
    d.lTheme = mkSec(IDS_L_THEME); d.lLang = mkSec(IDS_L_LANG); d.lLog = mkSec(IDS_L_LOG); d.lRefresh = mkSec(IDS_L_REFRESH);
    d.cTheme = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_C_THEME, g.hInst, nullptr);
    d.cLang = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP | WS_VSCROLL, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_C_LANG, g.hInst, nullptr);
    d.cLog = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_C_LOG, g.hInst, nullptr);
    d.eRefresh = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"", WS_CHILD | WS_VISIBLE | ES_NUMBER | WS_TABSTOP, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_E_REFRESH, g.hInst, nullptr);
    d.cFont = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_C_FONT, g.hInst, nullptr);
    d.cSize = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | WS_TABSTOP, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)IDS_C_SIZE, g.hInst, nullptr);
    d.lFont = mkSec(IDS_L_FONT);
    d.lSize = mkSec(IDS_L_SIZE);
    d.tTray = CreatePsqcToggle(d.hwnd, IDS_T1, -1, S().trayIcon);
    d.tMinTray = CreatePsqcToggle(d.hwnd, IDS_T2, -1, S().minimizeToTray);
    d.tStartMin = CreatePsqcToggle(d.hwnd, IDS_T3, -1, S().startMinimized);
    d.tNotify = CreatePsqcToggle(d.hwnd, IDS_T4, -1, S().notifications);
    d.tConfirm = CreatePsqcToggle(d.hwnd, IDS_T5, -1, S().confirmClearAll);
    d.tAutoUpd = CreatePsqcToggle(d.hwnd, IDS_T6, -1, S().autoUpdate);
    d.tElev = CreatePsqcToggle(d.hwnd, IDS_T7, -1, S().autoElevate);
    d.tBackup = CreatePsqcToggle(d.hwnd, IDS_T8, -1, S().backupBeforePurge);
    d.tAutoStart = CreatePsqcToggle(d.hwnd, IDS_T9, -1, AutostartEnabled());
    d.bLogs = CreatePsqcButton(d.hwnd, IDS_B_LOGS, STR_BTN_OPEN_LOGS, GLYPH_FOLDER, BK_Standard);
    d.bBackups = CreatePsqcButton(d.hwnd, IDS_B_BACKUPS, STR_BTN_OPEN_BACKUPS, GLYPH_FOLDER, BK_Standard);
    d.bRestore = CreatePsqcButton(d.hwnd, IDS_B_RESTORE, STR_BTN_RESTORE, GLYPH_BACK, BK_Standard);
    d.bUninstall = CreatePsqcButton(d.hwnd, IDS_B_UNINSTALL, STR_BTN_UNINSTALL, GLYPH_DELETE, BK_Danger);
    d.note = mkSec(IDS_NOTE);
    for (int i = 0; i < 9; ++i)
        d.tl[i] = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, d.hwnd, (HMENU)(INT_PTR)(IDS_T_TRAY + i), g.hInst, nullptr);

    ApplyDlgFonts();

    ApplyTexts();
    ThemeDarkControls(d.hwnd, Th().dark);
    util::ApplyRtlLayout(d.hwnd);

    int x = Sc(20), y = Sc(16);
    SetWindowPos(d.secApp, nullptr, x, y, W - 2 * x, Sc(20), SWP_NOZORDER); y += Sc(24);
    SetWindowPos(d.lTheme, nullptr, x, y + Sc(4), Sc(90), Sc(20), SWP_NOZORDER);
    SetWindowPos(d.cTheme, nullptr, x + Sc(94), y, Sc(140), Sc(24), SWP_NOZORDER);
    SetWindowPos(d.lLang, nullptr, x + Sc(250), y + Sc(4), Sc(60), Sc(20), SWP_NOZORDER);
    SetWindowPos(d.cLang, nullptr, x + Sc(314), y, Sc(160), Sc(200), SWP_NOZORDER);
    y += Sc(32);
    SetWindowPos(d.lFont, nullptr, x, y + Sc(4), Sc(90), Sc(20), SWP_NOZORDER);
    SetWindowPos(d.cFont, nullptr, x + Sc(94), y, Sc(200), Sc(24), SWP_NOZORDER);
    SetWindowPos(d.lSize, nullptr, x + Sc(306), y + Sc(4), Sc(110), Sc(20), SWP_NOZORDER);
    SetWindowPos(d.cSize, nullptr, x + Sc(420), y, Sc(54), Sc(24), SWP_NOZORDER);
    y += Sc(32);
    SetWindowPos(d.secBeh, nullptr, x, y, W - 2 * x, Sc(20), SWP_NOZORDER); y += Sc(24);
    HWND tgl[9] = { d.tTray, d.tMinTray, d.tStartMin, d.tNotify, d.tConfirm, d.tAutoUpd, d.tElev, d.tBackup, d.tAutoStart };
    int yy = y;
    for (int i = 0; i < 9; ++i) {
        SetWindowPos(tgl[i], nullptr, x, yy + Sc(1), Sc(44), Sc(22), SWP_NOZORDER);
        SetWindowPos(d.tl[i], nullptr, x + Sc(52), yy + Sc(4), W - x - Sc(72), Sc(20), SWP_NOZORDER);
        yy += Sc(28);
    }
    y = yy + Sc(6);
    SetWindowPos(d.secLog, nullptr, x, y, W - 2 * x, Sc(20), SWP_NOZORDER); y += Sc(24);
    auto TextCx = [&](const wchar_t* s) {
        HDC dc = GetDC(d.hwnd);
        HFONT of = (HFONT)SelectObject(dc, Th().font);
        SIZE ts{};
        GetTextExtentPoint32W(dc, s, (int)wcslen(s), &ts);
        SelectObject(dc, of);
        ReleaseDC(d.hwnd, dc);
        return ts.cx;
    };
    int wLog = TextCx(psqc::Str(STR_SET_LOGGING)) + Sc(10);
    SetWindowPos(d.lLog, nullptr, x, y + Sc(4), wLog, Sc(20), SWP_NOZORDER);
    SetWindowPos(d.cLog, nullptr, x + wLog + Sc(8), y, Sc(140), Sc(24), SWP_NOZORDER);
    y += Sc(34);
    int wRef = TextCx(psqc::Str(STR_SET_REFRESH)) + Sc(8);
    SetWindowPos(d.lRefresh, nullptr, x, y + Sc(4), wRef, Sc(20), SWP_NOZORDER);
    SetWindowPos(d.eRefresh, nullptr, x + wRef + Sc(8), y, Sc(60), Sc(24), SWP_NOZORDER);
    y += Sc(34);
    SetWindowPos(d.secData, nullptr, x, y, W - 2 * x, Sc(20), SWP_NOZORDER); y += Sc(26);
    // Кнопки данных — по ширине своего текста, «потоком» в две строки.
    {
        HDC dc = GetDC(d.hwnd);
        CtrlResizeToContent(d.bLogs, dc);
        CtrlResizeToContent(d.bBackups, dc);
        CtrlResizeToContent(d.bRestore, dc);
        CtrlResizeToContent(d.bUninstall, dc);
        ReleaseDC(d.hwnd, dc);
    }
    RECT br{};
    auto Bw = [&](HWND h) { GetWindowRect(h, &br); return br.right - br.left; };
    int wLogs = Bw(d.bLogs), wBack = Bw(d.bBackups), wRest = Bw(d.bRestore), wUnin = Bw(d.bUninstall);
    SetWindowPos(d.bLogs, nullptr, x, y, wLogs, Sc(32), SWP_NOZORDER);
    if (x + wLogs + Sc(8) + wBack <= W - x)
        SetWindowPos(d.bBackups, nullptr, x + wLogs + Sc(8), y, wBack, Sc(32), SWP_NOZORDER);
    else
        SetWindowPos(d.bBackups, nullptr, x, y + Sc(40), wBack, Sc(32), SWP_NOZORDER);
    y += Sc(40);
    SetWindowPos(d.bRestore, nullptr, x, y, wRest, Sc(32), SWP_NOZORDER);
    if (x + wRest + Sc(8) + wUnin <= W - x)
        SetWindowPos(d.bUninstall, nullptr, x + wRest + Sc(8), y, wUnin, Sc(32), SWP_NOZORDER);
    else
        SetWindowPos(d.bUninstall, nullptr, x, y + Sc(40), wUnin, Sc(32), SWP_NOZORDER);
    y += Sc(42);
    SetWindowPos(d.note, nullptr, x, y, W - 2 * x, Sc(18), SWP_NOZORDER);

    SetWindowTextW(d.eRefresh, std::to_wstring(S().refreshSec).c_str());
    util::CenterWindow(d.hwnd, g.hwndMain);
    ShowWindow(d.hwnd, SW_SHOW);
    UpdateWindow(d.hwnd);
#ifdef PSQC_QA_DELAY
    Sleep(700);
#endif
    util::QaShot(d.hwnd, L"settings");
    if (GetEnvironmentVariableW(L"PSQC_SHOTDIR", nullptr, 0) != 0)
        RestoreShow();   // QA: снять и диалог восстановления (раньше открывался всегда)
}
