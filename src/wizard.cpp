// Built-in installer / uninstaller / first-run wizard. Zero external dependencies.
#include "wizard.h"
#include "i18n.h"
#include "theme.h"
#include "settings.h"
#include "util.h"
#include "controls.h"

namespace psqc { const wchar_t* Str(int id); }

enum WizIds : int {
    IZW_ICON = 601, IZW_TITLE, IZW_TEXT,
    IZW_R_INSTALL, IZW_R_RUNONCE,
    IZW_R_ALL, IZW_R_CUR, IZW_SCOPE_NOTE,
    IZW_EDIT_DIR, IZW_BROWSE,
    IZW_C_DESK, IZW_C_START, IZW_C_AUTO, IZW_C_TRAY,
    IZW_STATUS,
    IZW_BACK, IZW_NEXT, IZW_CANCEL, IZW_RUNONCE_BTN,
};

struct WZ {
    HWND hwnd = nullptr;
    int page = 0;
    InstallOptions o;
    bool runOnceChoice = false;
    HWND stTitle, stText, rInstall;
    HWND rAll, rCur, scopeNote;
    HWND eDir, bBrowse;
    HWND cDesk, cStart, cAuto, cTray;
    HWND status;
    HWND bBack, bNext, bCancel, bRunOnce;
    WizardResult result = WR_Cancelled;
};
static WZ w;

std::wstring DefaultInstallDir(int scope) {
    if (scope == 1) {
        wchar_t buf[MAX_PATH]{};
        GetEnvironmentVariableW(L"ProgramFiles", buf, MAX_PATH);
        return util::JoinPath(buf, L"SpoolerCleanup");
    }
    PWSTR raw = nullptr;
    std::wstring base;
    if (SUCCEEDED(SHGetKnownFolderPath(FOLDERID_LocalAppData, 0, nullptr, &raw)) && raw) base = raw;
    if (raw) CoTaskMemFree(raw);
    return util::JoinPath(util::JoinPath(base, L"Programs"), L"SpoolerCleanup");
}

// ---- shortcuts / registry ---------------------------------------------------------
static bool MakeShortcut(const std::wstring& lnkPath, const std::wstring& target, const std::wstring& args) {
    size_t cut = lnkPath.find_last_of(L"\\/");
    if (cut != std::wstring::npos) util::EnsureDir(lnkPath.substr(0, cut));
    IShellLinkW* link = nullptr;
    HRESULT hr = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_IShellLinkW, (void**)&link);
    if (FAILED(hr)) return false;
    link->SetPath(target.c_str());
    if (!args.empty()) link->SetArguments(args.c_str());
    link->SetIconLocation(target.c_str(), 0);
    link->SetDescription(L"Print Spooler Queue Cleanup");
    IPersistFile* pf = nullptr;
    hr = link->QueryInterface(IID_IPersistFile, (void**)&pf);
    bool ok = false;
    if (SUCCEEDED(hr)) {
        ok = SUCCEEDED(pf->Save(lnkPath.c_str(), TRUE));
        pf->Release();
    }
    link->Release();
    return ok;
}

static std::wstring DesktopDir(bool allUsers) {
    PWSTR raw = nullptr;
    std::wstring out;
    KNOWNFOLDERID id = allUsers ? FOLDERID_PublicDesktop : FOLDERID_Desktop;
    if (SUCCEEDED(SHGetKnownFolderPath(id, 0, nullptr, &raw)) && raw) out = raw;
    if (raw) CoTaskMemFree(raw);
    return out;
}

static std::wstring ProgramsDir(bool allUsers) {
    PWSTR raw = nullptr;
    std::wstring out;
    KNOWNFOLDERID id = allUsers ? FOLDERID_CommonPrograms : FOLDERID_Programs;
    if (SUCCEEDED(SHGetKnownFolderPath(id, 0, nullptr, &raw)) && raw) out = raw;
    if (raw) CoTaskMemFree(raw);
    return out;
}

static std::wstring DesktopLnk(bool all) { return util::JoinPath(DesktopDir(all), L"Print Spooler Queue Cleanup.lnk"); }
static std::wstring ProgramsLnk(bool all) { return util::JoinPath(ProgramsDir(all), L"Print Spooler Queue Cleanup.lnk"); }

static const wchar_t* UNINSTALL_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\SpoolerCleanup";

bool InstallExecute(const InstallOptions& o, std::wstring& err) {
    if (!util::EnsureDir(o.dir)) { err = util::LastErrorText(GetLastError()); return false; }
    std::wstring target = util::JoinPath(o.dir, PSQC_EXE_NAME);
    if (!util::IEquals(util::ExePath(), target)) {
        if (!CopyFileW(util::ExePath().c_str(), target.c_str(), FALSE)) {
            err = util::LastErrorText(GetLastError());
            return false;
        }
    }
    bool all = o.scope == 1;
    if (o.desktop) MakeShortcut(DesktopLnk(all), target, L"");
    if (o.startMenu) MakeShortcut(ProgramsLnk(all), target, L"");

    S().installScope = o.scope;
    S().installPath = o.dir;
    S().firstRunDone = true;
    if (o.trayAtLogon) S().startMinimized = true;
    SettingsSave();

    if (o.autostart) {
        std::wstring aerr;
        AutostartSet(true, aerr);   // best effort; scope-aware
    }

    HKEY root = (all && util::IsElevated()) ? HKEY_LOCAL_MACHINE : HKEY_CURRENT_USER;
    util::RegSetString(root, UNINSTALL_KEY, L"DisplayName", PSQC_PRODUCT_NAME);
    util::RegSetString(root, UNINSTALL_KEY, L"DisplayVersion", PSQC_VERSION_STR_W);
    util::RegSetString(root, UNINSTALL_KEY, L"Publisher", PSQC_COMPANY);
    util::RegSetString(root, UNINSTALL_KEY, L"DisplayIcon", target.c_str());
    util::RegSetString(root, UNINSTALL_KEY, L"InstallLocation", o.dir.c_str());
    util::RegSetString(root, UNINSTALL_KEY, L"UninstallString",
                       (L"\"" + target + L"\" /uninstall").c_str());
    util::RegSetString(root, UNINSTALL_KEY, L"QuietUninstallString",
                       (L"\"" + target + L"\" /uninstall /silent").c_str());
    util::RegSetDword(root, UNINSTALL_KEY, L"EstimatedSize", 1500);
    util::RegSetDword(root, UNINSTALL_KEY, L"NoModify", 1);
    util::RegSetDword(root, UNINSTALL_KEY, L"NoRepair", 1);
    return true;
}

void UninstallStart(bool silent) {
    if (!silent) {
        std::wstring q = util::Fmt(psqc::Str(STR_UNINSTALL_Q), PSQC_PRODUCT_NAME);
        if (MessageBoxW(g.hwndMain ? g.hwndMain : nullptr, q.c_str(), psqc::Str(STR_APP_TITLE),
                        MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2) != IDYES)
            return;
    }
    bool all = S().installScope == 1;
    std::wstring err;
    AutostartSet(false, err);
    util::RemoveFile(DesktopLnk(all));
    util::RemoveFile(DesktopLnk(!all));
    util::RemoveFile(ProgramsLnk(all));
    util::RemoveFile(ProgramsLnk(!all));
    util::RegDeleteTree(HKEY_CURRENT_USER, UNINSTALL_KEY);
    util::RegDeleteTree(HKEY_LOCAL_MACHINE, UNINSTALL_KEY);
    // schedule removal of install dir (including this exe) after process exit
    std::wstring dir = S().installPath.empty() ? util::ExeDir() : S().installPath;
    std::wstring cmd = util::Fmt(L"/c timeout /t 2 /nobreak >nul & del /f /q \"%s\" & rd /s /q \"%s\"",
                                 util::JoinPath(dir, PSQC_EXE_NAME).c_str(), dir.c_str());
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC;
    sei.lpVerb = L"open";
    sei.lpFile = L"cmd.exe";
    sei.lpParameters = cmd.c_str();
    sei.nShow = SW_HIDE;
    ShellExecuteExW(&sei);
    if (!silent && g.hwndMain)
        MessageBoxW(g.hwndMain, psqc::Str(STR_UNINSTALL_DONE), psqc::Str(STR_APP_TITLE), MB_ICONINFORMATION);
    if (g.hwndMain) DestroyWindow(g.hwndMain);
    else PostQuitMessage(0);
}

int CliInstall(int argc, wchar_t** argv) {
    InstallOptions o;
    o.dir = DefaultInstallDir(0);
    bool silent = false;
    for (int i = 1; i < argc; ++i) {
        std::wstring a = argv[i];
        if (util::IEquals(a, L"/allusers")) { o.scope = 1; o.dir = DefaultInstallDir(1); }
        else if (util::IEquals(a, L"/currentuser")) { o.scope = 0; o.dir = DefaultInstallDir(0); }
        else if (util::StartsWithI(a, L"/dir=")) o.dir = a.substr(5);
        else if (util::IEquals(a, L"/nodesktop")) o.desktop = false;
        else if (util::IEquals(a, L"/nostartmenu")) o.startMenu = false;
        else if (util::IEquals(a, L"/noautostart")) o.autostart = false;
        else if (util::IEquals(a, L"/notray")) o.trayAtLogon = false;
        else if (util::IEquals(a, L"/silent")) silent = true;
    }
    if (o.scope == 1 && !util::IsElevated()) {
        DWORD code = 0;
        std::wstring params = util::Fmt(L"/install /allusers /silent /dir=\"%s\"", o.dir.c_str());
        if (!util::RunElevatedAndWait(util::ExePath().c_str(), params.c_str(), &code)) return 2;
        return (int)code;
    }
    std::wstring err;
    if (!InstallExecute(o, err)) {
        if (!silent) MessageBoxW(nullptr, util::Fmt(psqc::Str(STR_WIZ_ERR_COPY), err.c_str()).c_str(),
                                 psqc::Str(STR_WIZ_TITLE), MB_ICONERROR);
        return 1;
    }
    return 0;
}

int CliUninstall(bool silent) {
    UninstallStart(silent);
    return 0;
}

// ---- wizard UI -------------------------------------------------------------------
static void ShowPage(int p) {
    w.page = p;
    auto Show = [](HWND h, bool v) { ShowWindow(h, v ? SW_SHOW : SW_HIDE); };
    Show(w.rInstall, p == 0); Show(w.bRunOnce, p == 0);
    Show(w.stText, p == 0 || p == 4);   // pages 1-3 reuse the title string as body text

    Show(w.rAll, p == 1); Show(w.rCur, p == 1); Show(w.scopeNote, p == 1);
    Show(w.eDir, p == 2); Show(w.bBrowse, p == 2);
    Show(w.cDesk, p == 3); Show(w.cStart, p == 3); Show(w.cAuto, p == 3); Show(w.cTray, p == 3);
    Show(w.status, p == 4);
    Show(w.bBack, p > 0 && p < 4);
    Show(w.bCancel, true);
    EnableWindow(w.bNext, TRUE);   // page 4 = Finish, it must stay enabled
    CtrlSetStrId(w.bNext, p == 4 ? STR_WIZ_LAUNCH : STR_WIZ_NEXT);
    if (p == 4) SetWindowTextW(w.status, w.o.dir.c_str());
    std::wstring title;
    switch (p) {
        case 0: title = psqc::Str(STR_WIZ_WELCOME);
                SetWindowTextW(w.stText, util::Fmt(psqc::Str(STR_WIZ_WELCOME_TEXT), PSQC_PRODUCT_NAME).c_str());
                break;
        case 1: title = psqc::Str(STR_WIZ_SCOPE); SetWindowTextW(w.stText, psqc::Str(STR_WIZ_SCOPE)); break;
        case 2: title = psqc::Str(STR_WIZ_FOLDER); SetWindowTextW(w.stText, psqc::Str(STR_WIZ_FOLDER)); break;
        case 3: title = psqc::Str(STR_WIZ_OPTIONS); SetWindowTextW(w.stText, psqc::Str(STR_WIZ_OPTIONS)); break;
        case 4: title = psqc::Str(STR_WIZ_DONE);
                SetWindowTextW(w.stText, util::Fmt(psqc::Str(STR_WIZ_DONE_TEXT), PSQC_PRODUCT_NAME).c_str());
                break;
    }
    SetWindowTextW(w.stTitle, title.c_str());
    InvalidateRect(w.hwnd, nullptr, TRUE);
    UpdateWindow(w.hwnd);
    util::QaShot(w.hwnd, util::Fmt(L"wizard_p%d", p).c_str());
}

static void DoInstall() {
    w.o.scope = SendMessageW(w.rAll, BM_GETCHECK, 0, 0) == BST_CHECKED ? 1 : 0;
    wchar_t buf[MAX_PATH]{};
    GetWindowTextW(w.eDir, buf, MAX_PATH);
    w.o.dir = buf;
    w.o.desktop = SendMessageW(w.cDesk, BM_GETCHECK, 0, 0) == BST_CHECKED;
    w.o.startMenu = SendMessageW(w.cStart, BM_GETCHECK, 0, 0) == BST_CHECKED;
    w.o.autostart = SendMessageW(w.cAuto, BM_GETCHECK, 0, 0) == BST_CHECKED;
    w.o.trayAtLogon = SendMessageW(w.cTray, BM_GETCHECK, 0, 0) == BST_CHECKED;
    SetWindowTextW(w.status, psqc::Str(STR_WIZ_INSTALLING));
    UpdateWindow(w.hwnd);

    if (w.o.scope == 1 && !util::IsElevated()) {
        DWORD code = 0;
        std::wstring params = util::Fmt(L"/install /allusers /silent /dir=\"%s\"%s%s%s%s",
            w.o.dir.c_str(),
            w.o.desktop ? L"" : L" /nodesktop",
            w.o.startMenu ? L"" : L" /nostartmenu",
            w.o.autostart ? L"" : L" /noautostart",
            w.o.trayAtLogon ? L"" : L" /notray");
        if (!util::RunElevatedAndWait(util::ExePath().c_str(), params.c_str(), &code)) {
            MessageBoxW(w.hwnd, psqc::Str(STR_WIZ_ERR_ADMIN), psqc::Str(STR_WIZ_TITLE), MB_ICONERROR);
            ShowPage(1);
            return;
        }
        S().installScope = 1;
        S().installPath = w.o.dir;
        S().firstRunDone = true;
        SettingsSave();
    } else {
        std::wstring err;
        if (!InstallExecute(w.o, err)) {
            MessageBoxW(w.hwnd, util::Fmt(psqc::Str(STR_WIZ_ERR_COPY), err.c_str()).c_str(),
                        psqc::Str(STR_WIZ_TITLE), MB_ICONERROR);
            return;
        }
    }
    ShowPage(4);
}

static LRESULT CALLBACK WizardProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_ERASEBKGND: { RECT rc{}; GetClientRect(hwnd, &rc); FillRect((HDC)wp, &rc, Th().brBg); return 1; }
        case WM_CTLCOLORSTATIC: case WM_CTLCOLORBTN:
            return (LRESULT)ThemeCtlColorBg((HDC)wp);
        case WM_DRAWITEM:
            // Without this the wizard's owner-draw buttons (Run once / Back / Next /
            // Cancel) are never painted: BS_OWNERDRAW parents must draw them here.
            if (CtrlHandleDrawItem((DRAWITEMSTRUCT*)lp)) return TRUE;
            return 0;
        case WM_PAINT: {
            PAINTSTRUCT ps{};
            HDC hdc = BeginPaint(hwnd, &ps);
            int dpi = util::DpiOf(hwnd);
            HICON ico = util::LoadAppIcon(MulDiv(40, dpi, 96), MulDiv(40, dpi, 96));
            DrawIconEx(hdc, MulDiv(20, dpi, 96), MulDiv(18, dpi, 96), ico, MulDiv(40, dpi, 96), MulDiv(40, dpi, 96), 0, nullptr, DI_NORMAL);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_COMMAND: {
            int id = LOWORD(wp);
            if (HIWORD(wp) == BN_CLICKED) {
                switch (id) {
                    case IZW_CANCEL: w.result = WR_Cancelled; DestroyWindow(hwnd); return 0;
                    case IZW_RUNONCE_BTN: {
                        S().firstRunDone = true;
                        SettingsSave();
                        w.result = WR_RunOnce;
                        DestroyWindow(hwnd);
                        return 0;
                    }
                    case IZW_BACK: if (w.page > 0) ShowPage(w.page - 1); return 0;
                    case IZW_NEXT:
                        if (w.page == 0) { ShowPage(1); return 0; }
                        if (w.page == 1) { ShowPage(2); return 0; }
                        if (w.page == 2) { ShowPage(3); return 0; }
                        if (w.page == 3) { DoInstall(); return 0; }
                        if (w.page == 4) { w.result = WR_Installed; DestroyWindow(hwnd); return 0; }
                        return 0;
                    case IZW_BROWSE: {
                        IFileOpenDialog* dlg = nullptr;
                        if (SUCCEEDED(CoCreateInstance(CLSID_FileOpenDialog, nullptr, CLSCTX_INPROC_SERVER, IID_IFileOpenDialog, (void**)&dlg))) {
                            DWORD opts = 0;
                            dlg->GetOptions(&opts);
                            dlg->SetOptions(opts | FOS_PICKFOLDERS | FOS_FORCEFILESYSTEM);
                            if (SUCCEEDED(dlg->Show(hwnd))) {
                                IShellItem* item = nullptr;
                                if (SUCCEEDED(dlg->GetResult(&item))) {
                                    PWSTR path = nullptr;
                                    if (SUCCEEDED(item->GetDisplayName(SIGDN_FILESYSPATH, &path))) {
                                        SetWindowTextW(w.eDir, path);
                                        CoTaskMemFree(path);
                                    }
                                    item->Release();
                                }
                            }
                            dlg->Release();
                        }
                        return 0;
                    }
                }
            }
            return 0;
        }
        case WM_CLOSE: w.result = WR_Cancelled; DestroyWindow(hwnd); return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

WizardResult WizardRun() {
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = WizardProc;
    wc.hInstance = g.hInst;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = util::LoadAppIcon(32, 32);
    wc.lpszClassName = PSQC_WIZARD_CLASS;
    RegisterClassExW(&wc);
#ifdef PSQC_QA_DELAY
    OutputDebugStringA("PSQC: wiz register");
#endif

    int dpi = util::DpiOf(nullptr);
    int W = MulDiv(560, dpi, 96), H = MulDiv(430, dpi, 96);
    w = WZ{};
    w.o.dir = DefaultInstallDir(0);
    w.hwnd = CreateWindowExW(0, PSQC_WIZARD_CLASS, psqc::Str(STR_WIZ_TITLE),
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, W, H,
                             nullptr, nullptr, g.hInst, nullptr);
    g.hwndWizard = w.hwnd;
    util::SetClientSize(w.hwnd, W, H);   // W x H is the client area, not the frame
#ifdef PSQC_QA_DELAY
    OutputDebugStringA("PSQC: wiz top created"); Sleep(900);
#endif
    ThemeDarkTitlebar(w.hwnd, Th().dark);

    auto S = [&](int v) { return MulDiv(v, dpi, 96); };
    auto mk = [&](int id, const wchar_t* cls, DWORD style, const wchar_t* text) {
        return CreateWindowExW(0, cls, text, WS_CHILD | style, 0, 0, 10, 10, w.hwnd, (HMENU)(INT_PTR)id, g.hInst, nullptr);
    };
    w.stTitle = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX, S(72), S(20), W - S(92), S(30), w.hwnd, (HMENU)(INT_PTR)IZW_TITLE, g.hInst, nullptr);
    w.stText = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX, S(72), S(52), W - S(100), S(90), w.hwnd, (HMENU)(INT_PTR)IZW_TEXT, g.hInst, nullptr);
    w.rInstall = mk(IZW_R_INSTALL, L"BUTTON", WS_VISIBLE | BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, psqc::Str(STR_WIZ_INSTALL));
    w.rAll = mk(IZW_R_ALL, L"BUTTON", BS_AUTORADIOBUTTON | WS_GROUP | WS_TABSTOP, psqc::Str(STR_WIZ_SCOPE_ALL));
    w.rCur = mk(IZW_R_CUR, L"BUTTON", BS_AUTORADIOBUTTON | WS_TABSTOP, psqc::Str(STR_WIZ_SCOPE_CUR));
    w.scopeNote = mk(IZW_SCOPE_NOTE, L"STATIC", SS_LEFT | SS_NOPREFIX, psqc::Str(STR_WIZ_SCOPE_NOTE));
    w.eDir = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", w.o.dir.c_str(), WS_CHILD | ES_AUTOHSCROLL | WS_TABSTOP, 0, 0, 10, 10, w.hwnd, (HMENU)(INT_PTR)IZW_EDIT_DIR, g.hInst, nullptr);
    w.bBrowse = CreatePsqcButton(w.hwnd, IZW_BROWSE, STR_WIZ_BROWSE, GLYPH_FOLDER, BK_Standard);
    w.cDesk = mk(IZW_C_DESK, L"BUTTON", BS_AUTOCHECKBOX | WS_TABSTOP, psqc::Str(STR_WIZ_OPT_DESKTOP));
    w.cStart = mk(IZW_C_START, L"BUTTON", BS_AUTOCHECKBOX | WS_TABSTOP, psqc::Str(STR_WIZ_OPT_STARTMENU));
    w.cAuto = mk(IZW_C_AUTO, L"BUTTON", BS_AUTOCHECKBOX | WS_TABSTOP, psqc::Str(STR_WIZ_OPT_AUTOSTART));
    w.cTray = mk(IZW_C_TRAY, L"BUTTON", BS_AUTOCHECKBOX | WS_TABSTOP, psqc::Str(STR_WIZ_OPT_TRAY));
    w.status = mk(IZW_STATUS, L"STATIC", SS_LEFT | SS_NOPREFIX, L"");
    w.bBack = CreatePsqcButton(w.hwnd, IZW_BACK, STR_WIZ_BACK, GLYPH_BACK, BK_Standard);
    w.bNext = CreatePsqcButton(w.hwnd, IZW_NEXT, STR_WIZ_NEXT, 0, BK_Accent);
    w.bCancel = CreatePsqcButton(w.hwnd, IZW_CANCEL, STR_WIZ_CANCEL, 0, BK_Subtle);
    w.bRunOnce = CreatePsqcButton(w.hwnd, IZW_RUNONCE_BTN, STR_WIZ_RUN_ONCE, 0, BK_Standard);

    SendMessageW(w.stTitle, WM_SETFONT, (WPARAM)Th().fontBig, TRUE);
    SendMessageW(w.stText, WM_SETFONT, (WPARAM)Th().font, TRUE);
    for (HWND h : { w.rInstall, w.rAll, w.rCur, w.scopeNote, w.eDir, w.cDesk, w.cStart, w.cAuto, w.cTray, w.status })
        SendMessageW(h, WM_SETFONT, (WPARAM)Th().font, TRUE);
    OutputDebugStringA("PSQC: wiz children done");
    SendMessageW(w.rInstall, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(w.rCur, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(w.cDesk, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(w.cStart, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(w.cAuto, BM_SETCHECK, BST_CHECKED, 0);
    SendMessageW(w.cTray, BM_SETCHECK, BST_CHECKED, 0);
    ThemeDarkControls(w.hwnd, Th().dark);

    int y = S(150);
    SetWindowPos(w.rInstall, nullptr, S(72), y, W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.bRunOnce, nullptr, S(72), y + S(30), S(230), S(32), SWP_NOZORDER);
    SetWindowPos(w.rAll, nullptr, S(72), y, W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.rCur, nullptr, S(72), y + S(26), W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.scopeNote, nullptr, S(92), y + S(52), W - S(120), S(36), SWP_NOZORDER);
    SetWindowPos(w.eDir, nullptr, S(72), y, W - S(72) - S(110), S(26), SWP_NOZORDER);
    SetWindowPos(w.bBrowse, nullptr, W - S(120), y - S(3), S(100), S(32), SWP_NOZORDER);
    SetWindowPos(w.cDesk, nullptr, S(72), y, W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.cStart, nullptr, S(72), y + S(26), W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.cAuto, nullptr, S(72), y + S(52), W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.cTray, nullptr, S(72), y + S(78), W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.status, nullptr, S(72), y + S(20), W - S(100), S(22), SWP_NOZORDER);
    SetWindowPos(w.bBack, nullptr, W - S(230), H - S(56), S(100), S(32), SWP_NOZORDER);
    SetWindowPos(w.bNext, nullptr, W - S(122), H - S(56), S(100), S(32), SWP_NOZORDER);
    SetWindowPos(w.bCancel, nullptr, S(20), H - S(56), S(100), S(32), SWP_NOZORDER);

    ShowPage(0);
    util::ApplyRtlLayout(w.hwnd);
    util::CenterWindow(w.hwnd, nullptr);
    ShowWindow(w.hwnd, SW_SHOW);
    UpdateWindow(w.hwnd);
#ifdef PSQC_QA_DELAY
    OutputDebugStringA("PSQC: wiz shown"); Sleep(900);
#endif
    if (getenv("PSQC_SHOTDIR")) {          // QA self-capture: walk all pages, then run-once
        for (int p = 0; p <= 4; ++p) ShowPage(p);
        DestroyWindow(w.hwnd);
        g.hwndWizard = nullptr;
        return WR_RunOnce;
    }

    MSG msg;
    // The window check must gate EVERY iteration, not run after DispatchMessageW:
    // a mouse click is delivered through IsDialogMessageW, which runs the wizard command,
    // destroys the window inside the handler and still reports the message as handled.
    // With the check at the bottom only, that 'continue' skipped it and the loop waited in
    // GetMessageW forever - the process stayed alive, windowless, holding the single-instance
    // mutex, so every later start exited silently (measured on Windows 11).
    while (IsWindow(w.hwnd) && GetMessageW(&msg, nullptr, 0, 0) > 0) {
        if (IsDialogMessageW(w.hwnd, &msg)) continue;
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }
    g.hwndWizard = nullptr;
    return w.result;
}
