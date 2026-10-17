// Main window implementation.
#include "main_window.h"
#include "controls.h"
#include "theme.h"
#include "i18n.h"
#include "settings.h"
#include "spooler.h"
#include "util.h"
#include "tray.h"
#include "about.h"
#include "settings_dialog.h"
#include "update.h"

void TrayHandleMessage(WPARAM wp, LPARAM lp);   // tray.cpp

namespace psqc { const wchar_t* Str(int id); }

struct MW {
    HWND hwnd = nullptr;
    HWND list = nullptr;
    HWND btnRefresh = nullptr, btnDelSel = nullptr, btnClearAll = nullptr, btnPause = nullptr, btnRestart = nullptr;
    HWND btnSettings = nullptr, btnTheme = nullptr, btnAbout = nullptr;
    HWND emptyT = nullptr, emptyH = nullptr;
    HWND cmbLang = nullptr, cmbTheme = nullptr, tglTray = nullptr, tglAuto = nullptr;
    HWND lblLang = nullptr, lblTheme = nullptr, lblTray = nullptr, lblAuto = nullptr;
    HWND chipSvc = nullptr, chipUser = nullptr, chipPortable = nullptr;
    HWND tips = nullptr;   // общий tooltip: полные подписи переключателей
    std::vector<spooler::PrinterJobs> snapshot;
    std::vector<spooler::JobInfo> flat;
    bool busy = false;
    bool anyPaused = false;
    int printersWithJobs = 0;
    SYSTEMTIME lastRefresh{};
};
static MW m;

HWND MainWindowHandle() { return m.hwnd; }
bool MainWindowBusy() { return m.busy; }

// ---- tooltips: полная надпись при наведении -----------------------------------
static void TipsEnsure(HWND parent) {
    if (m.tips || !parent) return;
    m.tips = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                             WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP, 0, 0, 0, 0,
                             parent, nullptr, g.hInst, nullptr);
    if (m.tips) SetWindowPos(m.tips, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
}
static void TipTool(HWND target, int strId, bool add) {
    if (!m.tips || !target) return;
    TOOLINFOW ti{};
    ti.cbSize = sizeof(ti);
    ti.uFlags = TTF_SUBCLASS | TTF_IDISHWND;
    ti.hwnd = m.hwnd;
    ti.uId = (UINT_PTR)target;
    ti.lpszText = (LPWSTR)psqc::Str(strId);
    SendMessageW(m.tips, add ? TTM_ADDTOOLW : TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti);
}
static void TipsAdd() {
    TipsEnsure(m.hwnd);
    TipTool(m.tglTray, STR_SET_TRAY, true);
    TipTool(m.lblTray, STR_SET_TRAY, true);
    TipTool(m.tglAuto, STR_SET_AUTOSTART, true);
    TipTool(m.lblAuto, STR_SET_AUTOSTART, true);
}
static void TipsUpdateTexts() {
    TipTool(m.tglTray, STR_SET_TRAY, false);
    TipTool(m.lblTray, STR_SET_TRAY, false);
    TipTool(m.tglAuto, STR_SET_AUTOSTART, false);
    TipTool(m.lblAuto, STR_SET_AUTOSTART, false);
}

// ---- layout ------------------------------------------------------------------
static int g_minCxClient = 0;   // минимальная ширина клиента (считается в Layout)

static void Layout() {
    if (!m.hwnd) return;
    static bool s_inLayout = false;
    if (s_inLayout) return;   // авто-подгон ширины внутри Layout вызывает WM_SIZE
    s_inLayout = true;
    RECT rc{};
    GetClientRect(m.hwnd, &rc);
    int dpi = util::DpiOf(m.hwnd);
    auto S = [&](int v) { return MulDiv(v, dpi, 96); };
    int margin = S(16);
    int headerH = S(56);
    int toolH = S(36);
    int setRowH = S(34);
    int statusH = S(26);
    int y = headerH + S(6);

    // toolbar buttons: left group then right group
    auto Place = [&](HWND h, int& x, int w) {
        SetWindowPos(h, nullptr, x, y, w, toolH, SWP_NOZORDER);
        x += w + S(8);
    };
    int x = margin;
    HDC hdc = GetDC(m.hwnd);
    CtrlResizeToContent(m.btnRefresh, hdc);
    CtrlResizeToContent(m.btnDelSel, hdc);
    CtrlResizeToContent(m.btnClearAll, hdc);
    CtrlResizeToContent(m.btnPause, hdc);
    CtrlResizeToContent(m.btnRestart, hdc);
    CtrlResizeToContent(m.btnSettings, hdc);
    CtrlResizeToContent(m.btnTheme, hdc);
    CtrlResizeToContent(m.btnAbout, hdc);
    ReleaseDC(m.hwnd, hdc);
    RECT br{};
    auto W = [&](HWND h) { GetWindowRect(h, &br); return br.right - br.left; };
    Place(m.btnRefresh, x, W(m.btnRefresh));
    Place(m.btnDelSel, x, W(m.btnDelSel));
    Place(m.btnClearAll, x, W(m.btnClearAll));
    x += S(6);
    Place(m.btnPause, x, W(m.btnPause));
    Place(m.btnRestart, x, W(m.btnRestart));
    int xr = rc.right - margin;
    if (xr < x + S(120)) xr = x + S(120);   // narrow-window clamp: never overlap left group
    xr -= W(m.btnAbout); SetWindowPos(m.btnAbout, nullptr, xr, y, W(m.btnAbout), toolH, SWP_NOZORDER);
    xr -= S(8) + W(m.btnTheme); SetWindowPos(m.btnTheme, nullptr, xr, y, W(m.btnTheme), toolH, SWP_NOZORDER);
    xr -= S(8) + W(m.btnSettings); SetWindowPos(m.btnSettings, nullptr, xr, y, W(m.btnSettings), toolH, SWP_NOZORDER);

    y += toolH + S(10);
    int listH = rc.bottom - statusH - setRowH - S(18) - y;
    if (listH < S(80)) listH = S(80);
    SetWindowPos(m.list, nullptr, margin, y, rc.right - margin * 2, listH, SWP_NOZORDER);
    // Оверлей пустой очереди — ВНУТРИ рамки списка: фон STATIC перекрывал окантовку.
    int ovx = margin + S(6), ovw = rc.right - margin * 2 - S(12);
    SetWindowPos(m.emptyT, nullptr, ovx, y + listH / 2 - S(28), ovw, S(26), SWP_NOZORDER);
    SetWindowPos(m.emptyH, nullptr, ovx, y + listH / 2 + S(2), ovw, S(20), SWP_NOZORDER);

    int ys = rc.bottom - statusH - setRowH - S(8);
    // Ширина подписей — по фактическому тексту текущего языка и шрифта:
    // «Запускать вместе с Windows» больше не обрезается.
    auto TextCx = [&](const wchar_t* s) {
        HDC dc = GetDC(m.hwnd);
        HFONT of = (HFONT)SelectObject(dc, Th().font);
        SIZE ts{};
        GetTextExtentPoint32W(dc, s, (int)wcslen(s), &ts);
        SelectObject(dc, of);
        ReleaseDC(m.hwnd, dc);
        return ts.cx;
    };
    int wTray = TextCx(psqc::Str(STR_SET_TRAY)) + S(8);
    int wAuto = TextCx(psqc::Str(STR_SET_AUTOSTART)) + S(8);
    int cx = margin;
    SetWindowPos(m.lblLang, nullptr, cx, ys + S(6), S(70), S(20), SWP_NOZORDER);
    cx += S(74);
    SetWindowPos(m.cmbLang, nullptr, cx, ys + S(2), S(150), S(24), SWP_NOZORDER);
    cx += S(158);
    SetWindowPos(m.lblTheme, nullptr, cx, ys + S(6), S(50), S(20), SWP_NOZORDER);
    cx += S(54);
    SetWindowPos(m.cmbTheme, nullptr, cx, ys + S(2), S(110), S(24), SWP_NOZORDER);
    cx += S(118) + S(10);
    SetWindowPos(m.tglTray, nullptr, cx, ys + S(5), S(44), S(22), SWP_NOZORDER);
    cx += S(48);
    SetWindowPos(m.lblTray, nullptr, cx, ys + S(6), wTray, S(20), SWP_NOZORDER);
    cx += wTray + S(6);
    SetWindowPos(m.tglAuto, nullptr, cx, ys + S(5), S(44), S(22), SWP_NOZORDER);
    cx += S(48);
    SetWindowPos(m.lblAuto, nullptr, cx, ys + S(6), wAuto, S(20), SWP_NOZORDER);

    // header chips (right aligned in header band)
    int yc = S(14);
    int xc = rc.right - margin;
    auto PlaceChip = [&](HWND h, bool show) {
        if (!show) { ShowWindow(h, SW_HIDE); return; }
        ShowWindow(h, SW_SHOW);
        HDC dc = GetDC(h);
        CtrlResizeToContent(h, dc);
        ReleaseDC(h, dc);
        RECT cr{}; GetWindowRect(h, &cr);
        int w = cr.right - cr.left;
        xc -= w;
        SetWindowPos(h, nullptr, xc, yc, w, S(26), SWP_NOZORDER);
        xc -= S(8);
    };
    PlaceChip(m.chipPortable, g.portable);
    PlaceChip(m.chipUser, true);
    PlaceChip(m.chipSvc, true);

    // Минимальная ширина клиента = тулбар целиком (левая группа + правая).
    // Узкое окно раньше выталкивало правую группу кнопок за край.
    {
        int rightW = W(m.btnSettings) + S(8) + W(m.btnTheme) + S(8) + W(m.btnAbout);
        g_minCxClient = margin * 2 + (x - margin) + S(8) + rightW;
        RECT rw{};
        GetWindowRect(m.hwnd, &rw);
        RECT fr{ 0, 0, g_minCxClient, 400 };
        AdjustWindowRectEx(&fr, WS_OVERLAPPEDWINDOW, FALSE, 0);
        int needWin = fr.right - fr.left;
        if ((rw.right - rw.left) < needWin)
            SetWindowPos(m.hwnd, nullptr, 0, 0, needWin, rw.bottom - rw.top, SWP_NOMOVE | SWP_NOZORDER);
    }
    s_inLayout = false;

    // list columns scale
    int cw[5] = { S(260), S(130), S(90), S(130), S(150) };
    for (int i = 0; i < 5; ++i) ListView_SetColumnWidth(m.list, i, cw[i]);
}

// ---- list --------------------------------------------------------------------
static void SetupColumns() {
    LVCOLUMNW c{};
    c.mask = LVCF_TEXT | LVCF_WIDTH | LVCF_FMT;
    c.fmt = LVCFMT_LEFT;
    const int ids[5] = { STR_COL_DOCUMENT, STR_COL_OWNER, STR_COL_SIZE, STR_COL_STATUS, STR_COL_SUBMITTED };
    int widths[5] = { 260, 130, 90, 130, 150 };
    static bool created = false;
    for (int i = 0; i < 5; ++i) {
        c.pszText = (LPWSTR)psqc::Str(ids[i]);
        c.cx = widths[i];
        c.iSubItem = i;
        if (created) ListView_SetColumn(m.list, i, &c);
        else ListView_InsertColumn(m.list, i, &c);
    }
    created = true;
}

static void UpdateEmptyOverlay() {
    bool empty = m.flat.empty();
    ShowWindow(m.emptyT, empty ? SW_SHOW : SW_HIDE);
    ShowWindow(m.emptyH, empty ? SW_SHOW : SW_HIDE);
    ShowWindow(m.list, SW_SHOW);
}

static void UpdateHeaderChips() {
    DWORD st = spooler::ServiceState();
    const wchar_t* stStr = psqc::Str(STR_SVC_STOPPED);
    COLORREF dot = Th().err;
    if (st == SERVICE_RUNNING) { stStr = psqc::Str(STR_SVC_RUNNING); dot = Th().ok; }
    else if (st == SERVICE_START_PENDING || st == SERVICE_STOP_PENDING || st == SERVICE_CONTINUE_PENDING) {
        stStr = psqc::Str(STR_SVC_PENDING); dot = Th().warn;
    }
    CtrlSetText(m.chipSvc, util::Fmt(psqc::Str(STR_ST_SERVICE), stStr));
    CtrlSetDotColor(m.chipSvc, dot);
    CtrlSetText(m.chipUser, psqc::Str(g.elevated ? STR_ST_ADMIN : STR_ST_USER));
    CtrlSetDotColor(m.chipUser, g.elevated ? Th().warn : Th().subtext);
    CtrlSetText(m.chipPortable, psqc::Str(STR_PORTABLE_BADGE));
    CtrlSetDotColor(m.chipPortable, Th().accent);
}

void MainWindowSetJobCount(int n) { g.liveJobCount = n; }

static void RebuildList() {
    m.flat.clear();
    m.printersWithJobs = 0;
    for (auto& p : m.snapshot) if (!p.jobs.empty()) ++m.printersWithJobs;
    for (auto& p : m.snapshot) for (auto& j : p.jobs) m.flat.push_back(j);
    m.anyPaused = spooler::AnyPaused(m.snapshot);

    SendMessageW(m.list, WM_SETREDRAW, FALSE, 0);
    ListView_DeleteAllItems(m.list);
    ListView_RemoveAllGroups(m.list);
    ListView_EnableGroupView(m.list, TRUE);
    int gi = 0;
    for (auto& p : m.snapshot) {
        if (p.jobs.empty()) continue;
        LVGROUP grp{};
        grp.cbSize = sizeof(grp);
        grp.mask = LVGF_HEADER | LVGF_STATE;
        std::wstring hdr = util::Fmt(L"%s  (%zu)", p.name.c_str(), p.jobs.size());
        grp.pszHeader = (LPWSTR)hdr.c_str();
        grp.state = grp.stateMask = LVGS_NORMAL;
        int gid = ListView_InsertGroup(m.list, gi, &grp);
        for (auto& j : p.jobs) {
            LVITEMW it{};
            it.mask = LVIF_TEXT | LVIF_GROUPID | LVIF_PARAM;
            int idx = (int)m.flat.size();
            for (int k = 0; k < (int)m.flat.size(); ++k)
                if (m.flat[k].jobId == j.jobId && m.flat[k].printer == j.printer) { idx = k; break; }
            it.iItem = ListView_GetItemCount(m.list);
            it.iSubItem = 0;
            it.lParam = idx;
            it.iGroupId = gid;
            it.pszText = (LPWSTR)j.document.c_str();
            int row = ListView_InsertItem(m.list, &it);
            ListView_SetItemText(m.list, row, 1, (LPWSTR)j.owner.c_str());
            auto sz = util::FormatBytes(j.size);
            ListView_SetItemText(m.list, row, 2, (LPWSTR)sz.c_str());
            auto stt = psqc::Str(spooler::JobStatusStrId(j.status));
            ListView_SetItemText(m.list, row, 3, (LPWSTR)stt);
            auto tm = util::FormatSystemTimeUser(j.submitted);
            ListView_SetItemText(m.list, row, 4, (LPWSTR)tm.c_str());
        }
        ++gi;
    }
    SendMessageW(m.list, WM_SETREDRAW, TRUE, 0);
    UpdateEmptyOverlay();
    GetLocalTime(&m.lastRefresh);
    UpdateHeaderChips();
    MainWindowSetJobCount((int)m.flat.size());
    TrayUpdateCounts((int)m.flat.size());
    InvalidateRect(m.hwnd, nullptr, FALSE);   // status bar text
}

void MainWindowRefreshNow() {
    if (!m.hwnd) return;
    spooler::EnumerateAll(m.snapshot);
    RebuildList();
}

// ---- workers -------------------------------------------------------------------
static void SetBusy(bool b) {
    m.busy = b;
    for (HWND h : { m.btnRefresh, m.btnDelSel, m.btnClearAll, m.btnPause, m.btnRestart })
        CtrlSetEnabledVisual(h, !b);
}

void MainWindowBeginPurge(bool confirm) {
    if (m.busy) return;
    if (confirm && S().confirmClearAll) {
        int r = MessageBoxW(m.hwnd, psqc::Str(STR_MSG_CONFIRM_ALL), psqc::Str(STR_APP_TITLE),
                            MB_YESNO | MB_ICONWARNING | MB_DEFBUTTON2);
        if (r != IDYES) return;
    }
    SetBusy(true);
    std::thread([] {
        auto* res = new spooler::PurgeResult();
        spooler::PurgeAll(S().backupBeforePurge, *res);
        PostMessageW(m.hwnd, WM_PSQC_WORK_DONE, TASK_PURGE_ALL, (LPARAM)res);
    }).detach();
}

static void DeleteSelected() {
    if (m.busy) return;
    int sel = ListView_GetSelectedCount(m.list);
    if (sel == 0) {
        MessageBoxW(m.hwnd, psqc::Str(STR_MSG_SEL_NONE), psqc::Str(STR_APP_TITLE), MB_ICONINFORMATION);
        return;
    }
    std::vector<int> indices;
    int idx = -1;
    while ((idx = ListView_GetNextItem(m.list, idx, LVNI_SELECTED)) != -1) indices.push_back(idx);
    std::vector<spooler::JobInfo> jobs;
    for (int i : indices) {
        LVITEMW it{};
        it.iItem = i;
        it.mask = LVIF_PARAM;
        ListView_GetItem(m.list, &it);
        jobs.push_back(m.flat[it.lParam]);
    }
    int failed = 0;
    DWORD err = 0;
    for (auto& j : jobs)
        if (!spooler::DeleteJob(j.printer, j.jobId, err)) ++failed;
    MainWindowRefreshNow();
    if (failed) {
        if (err == ERROR_ACCESS_DENIED && !g.elevated && S().autoElevate) {
            if (MessageBoxW(m.hwnd, psqc::Str(STR_MSG_ELEVATE_Q), psqc::Str(STR_APP_TITLE),
                            MB_YESNO | MB_ICONWARNING) == IDYES) {
                DWORD code = 0;
                util::RunElevatedAndWait(util::ExePath().c_str(), L"/purge", &code);
                MainWindowRefreshNow();
            }
        } else {
            MessageBoxW(m.hwnd, util::Fmt(psqc::Str(STR_MSG_CLEAR_FAIL), util::LastErrorText(err).c_str()).c_str(),
                        psqc::Str(STR_APP_TITLE), MB_ICONERROR);
        }
    }
}

static void TogglePause() {
    if (m.busy) return;
    bool nowPause = !m.anyPaused;
    spooler::PauseAll(nowPause);
    MainWindowRefreshNow();
    CtrlSetStrId(m.btnPause, m.anyPaused ? STR_TB_RESUME : STR_TB_PAUSE);
    CtrlSetGlyph(m.btnPause, m.anyPaused ? GLYPH_PLAY : GLYPH_PAUSE);
}

static void RestartSpoolerAsync() {
    if (m.busy) return;
    SetBusy(true);
    std::thread([] {
        auto* res = new spooler::PurgeResult();
        std::wstring err;
        res->ok = spooler::RestartService(err);
        res->error = err;
        PostMessageW(m.hwnd, WM_PSQC_WORK_DONE, TASK_RESTART_SPOOLER, (LPARAM)res);
    }).detach();
}

// ---- theme / lang ----------------------------------------------------------------
static void FillThemeCombo() {
    SendMessageW(m.cmbTheme, CB_RESETCONTENT, 0, 0);
    SendMessageW(m.cmbTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_SYSTEM));
    SendMessageW(m.cmbTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_LIGHT));
    SendMessageW(m.cmbTheme, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_DARK));
    SendMessageW(m.cmbTheme, CB_SETCURSEL, S().themeMode, 0);
}

static void FillLangCombo() {
    SendMessageW(m.cmbLang, CB_RESETCONTENT, 0, 0);
    SendMessageW(m.cmbLang, CB_ADDSTRING, 0, (LPARAM)psqc::Str(STR_THEME_SYSTEM));
    for (int i = 0; i < LangCount(); ++i)
        SendMessageW(m.cmbLang, CB_ADDSTRING, 0, (LPARAM)Langs()[i].native);
    int cur = 0;
    if (!S().lang.empty()) {
        for (int i = 0; i < LangCount(); ++i)
            if (util::IEquals(Langs()[i].tag, S().lang)) { cur = i + 1; break; }
    }
    SendMessageW(m.cmbLang, CB_SETCURSEL, cur, 0);
}

static void ApplyStaticTexts() {
    SetWindowTextW(m.hwnd, psqc::Str(STR_APP_TITLE));
    SetWindowTextW(m.lblLang, psqc::Str(STR_SET_LANGUAGE));
    SetWindowTextW(m.lblTheme, psqc::Str(STR_SET_THEME));
    SetWindowTextW(m.lblTray, psqc::Str(STR_SET_TRAY));
    SetWindowTextW(m.lblAuto, psqc::Str(STR_SET_AUTOSTART));
    SetWindowTextW(m.emptyT, psqc::Str(STR_EMPTY_TITLE));
    SetWindowTextW(m.emptyH, psqc::Str(STR_EMPTY_HINT));
    SetupColumns();
    FillLangCombo();
    FillThemeCombo();
    CtrlRelabelAll();
    UpdateHeaderChips();
    TipsUpdateTexts();
}

static void ApplyFonts() {
    if (!m.hwnd) return;
    SendMessageW(m.list, WM_SETFONT, (WPARAM)Th().font, TRUE);
    SendMessageW(m.emptyT, WM_SETFONT, (WPARAM)Th().fontBig, TRUE);
    SendMessageW(m.emptyH, WM_SETFONT, (WPARAM)Th().font, TRUE);
    for (HWND h : { m.lblLang, m.lblTheme, m.lblTray, m.lblAuto, m.cmbLang, m.cmbTheme })
        SendMessageW(h, WM_SETFONT, (WPARAM)Th().font, TRUE);
}

void MainWindowReskin() {
    if (!m.hwnd) return;
    ThemeDarkTitlebar(m.hwnd, Th().dark);
    ThemeDarkControls(m.hwnd, Th().dark);
    ListView_SetBkColor(m.list, Th().surface2);
    ListView_SetTextBkColor(m.list, Th().surface2);
    ListView_SetTextColor(m.list, Th().text);
    ApplyStaticTexts();
    ApplyFonts();
    Layout();                       // ширины подписей зависят от языка и шрифта
    InvalidateRect(m.hwnd, nullptr, TRUE);
}

// ---- paint -----------------------------------------------------------------------
static void PaintHeader(HDC hdc, const RECT& rc) {
    int dpi = util::DpiOf(m.hwnd);
    auto S = [&](int v) { return MulDiv(v, dpi, 96); };
    HICON ico = util::LoadAppIcon(S(32), S(32));
    DrawIconEx(hdc, S(16), S(12), ico, S(32), S(32), 0, nullptr, DI_NORMAL);
    SetBkMode(hdc, TRANSPARENT);
    HFONT of = (HFONT)SelectObject(hdc, Th().fontBig);
    SetTextColor(hdc, Th().text);
    RECT tr{ S(56), S(8), rc.right - S(300), S(34) };
    DrawTextW(hdc, psqc::Str(STR_APP_TITLE), -1, &tr, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, Th().fontSmall);
    SetTextColor(hdc, Th().subtext);
    RECT tr2{ S(56), S(32), rc.right - S(300), S(52) };
    DrawTextW(hdc, psqc::Str(STR_APP_TAGLINE), -1, &tr2, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);
}

static void PaintStatus(HDC hdc, const RECT& rc) {
    int dpi = util::DpiOf(m.hwnd);
    auto S = [&](int v) { return MulDiv(v, dpi, 96); };
    int h = S(26);
    RECT band{ 0, rc.bottom - h, rc.right, rc.bottom };
    FillRect(hdc, &band, Th().brSurface);
    RECT line{ 0, rc.bottom - h, rc.right, rc.bottom - h + 1 };
    FillRect(hdc, &line, Th().brBg);
    SetBkMode(hdc, TRANSPARENT);
    HFONT of = (HFONT)SelectObject(hdc, Th().fontSmall);
    SetTextColor(hdc, Th().subtext);
    std::wstring left = util::Fmt(L"%s   ·   %s",
        util::Fmt(psqc::Str(STR_ST_JOBS), util::FormatInt(m.flat.size()).c_str()).c_str(),
        util::Fmt(psqc::Str(STR_ST_PRINTERS), util::FormatInt(m.printersWithJobs).c_str()).c_str());
    RECT lr{ S(16), band.top, rc.right - S(260), band.bottom };
    DrawTextW(hdc, left.c_str(), -1, &lr, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    std::wstring right = util::Fmt(psqc::Str(STR_ST_UPDATED), util::FormatSystemTimeUser(m.lastRefresh).c_str());
    RECT rr{ rc.right - S(250), band.top, rc.right - S(16), band.bottom };
    DrawTextW(hdc, right.c_str(), -1, &rr, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, of);
}

// ---- шапка таблицы: собственная отрисовка (subclass) ---------------------------------
// NM_CUSTOMDRAW шапки внутри ListView стабильно до родителя не доходит —
// рисуем шапку сами в WM_PAINT: выделенный фон, белый текст, разделители.
static LRESULT CALLBACK HeaderProc(HWND h, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR) {
    switch (msg) {
        case WM_ERASEBKGND:
            return 1;
        case WM_PAINT: {
            PAINTSTRUCT ps{};
            HDC hdc = BeginPaint(h, &ps);
            RECT rcTotal{};
            GetClientRect(h, &rcTotal);
            FillRect(hdc, &rcTotal, Th().brHeader);
            SetBkMode(hdc, TRANSPARENT);
            SetTextColor(hdc, Th().text);
            HFONT of = (HFONT)SelectObject(hdc, Th().font);
            UINT base = CurrentLangIsRtl() ? (UINT)(DT_RIGHT | DT_RTLREADING) : (UINT)DT_LEFT;
            int n = (int)SendMessageW(h, HDM_GETITEMCOUNT, 0, 0);
            for (int i = 0; i < n; ++i) {
                RECT rc{};
                if (!Header_GetItemRect(h, i, &rc)) continue;
                wchar_t txt[256]{};
                HDITEMW it{};
                it.mask = HDI_TEXT;
                it.pszText = txt;
                it.cchTextMax = 256;
                Header_GetItem(h, i, &it);
                RECT tr = rc;
                tr.left += util::Scale(h, 8);
                tr.right -= util::Scale(h, 8);
                DrawTextW(hdc, txt, -1, &tr, base | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
                RECT vln{ rc.right - 1, rc.top + util::Scale(h, 6), rc.right, rc.bottom - util::Scale(h, 6) };
                FillRect(hdc, &vln, Th().brBorder);
            }
            RECT ln{ rcTotal.left, rcTotal.bottom - 1, rcTotal.right, rcTotal.bottom };
            FillRect(hdc, &ln, Th().brBorder);
            SelectObject(hdc, of);
            EndPaint(h, &ps);
            return 0;
        }
    }
    return DefSubclassProc(h, msg, wp, lp);
}

// ---- wndproc ------------------------------------------------------------------------
static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_CREATE:
            return 0;
        case WM_ERASEBKGND: {
            RECT rc{};
            GetClientRect(hwnd, &rc);
            FillRect((HDC)wp, &rc, Th().brBg);
            return 1;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps{};
            HDC hdc = BeginPaint(hwnd, &ps);
            RECT rc{};
            GetClientRect(hwnd, &rc);
            PaintHeader(hdc, rc);
            PaintStatus(hdc, rc);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_CTLCOLORSTATIC: {
            HWND c = (HWND)lp;
            if (c == m.emptyT || c == m.emptyH) {
                // Оверлей пустой очереди лежит НА списке: фон — как у списка.
                SetBkColor((HDC)wp, Th().surface2);
                SetTextColor((HDC)wp, c == m.emptyT ? Th().text : Th().subtext);
                return (LRESULT)Th().brSurface2;
            }
            return (LRESULT)ThemeCtlColorBg((HDC)wp);
        }
        case WM_DRAWITEM:
            if (CtrlHandleDrawItem((DRAWITEMSTRUCT*)lp)) return TRUE;
            return 0;
        case WM_GETMINMAXINFO: {
            DefWindowProcW(hwnd, msg, wp, lp);
            if (g_minCxClient > 0) {
                auto* mmi = (MINMAXINFO*)lp;
                RECT fr{ 0, 0, g_minCxClient, 400 };
                AdjustWindowRectEx(&fr, WS_OVERLAPPEDWINDOW, FALSE, 0);
                mmi->ptMinTrackSize.x = fr.right - fr.left;
                int minH = MulDiv(420, util::DpiOf(hwnd), 96);
                if (mmi->ptMinTrackSize.y < minH) mmi->ptMinTrackSize.y = minH;
            }
            return 0;
        }
        case WM_SIZE:
            Layout();
            return 0;
        case WM_DPICHANGED:
            ThemeRefreshDpi(LOWORD(wp));
            Layout();
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case WM_SETTINGCHANGE:
            if (lp && wcscmp((const wchar_t*)lp, L"ImmersiveColorSet") == 0 && S().themeMode == THEME_SYSTEM) {
                ThemeSetMode(THEME_SYSTEM);
                MainWindowReskin();
                TrayReskin();
            }
            return 0;
        case WM_COMMAND: {
            int id = LOWORD(wp);
            int code = HIWORD(wp);
            if (code == BN_CLICKED || code == BN_DBLCLK) {
                switch (id) {
                    case IDC_BTN_REFRESH: MainWindowRefreshNow(); return 0;
                    case IDC_BTN_DELSEL: DeleteSelected(); return 0;
                    case IDC_BTN_CLEARALL: MainWindowBeginPurge(true); return 0;
                    case IDC_BTN_PAUSE: TogglePause(); return 0;
                    case IDC_BTN_RESTART: RestartSpoolerAsync(); return 0;
                    case IDC_BTN_SETTINGS: SettingsShow(); return 0;
                    case IDC_BTN_THEME: {
                        // Прямое переключение светлая<->тёмная по ФАКТИЧЕСКОЙ теме:
                        // прежний цикл трёх состояний давал «холостой» клик через «Как в системе».
                        int next = ThemeIsDarkEffective() ? THEME_LIGHT : THEME_DARK;
                        S().themeMode = next;
                        SettingsSave();
                        ThemeSetMode(next);
                        MainWindowReskin();
                        TrayReskin();
                        FillThemeCombo();
                        CtrlSetGlyph(m.btnTheme, Th().dark ? GLYPH_SUN : GLYPH_MOON);
                        return 0;
                    }
                    case IDC_BTN_ABOUT: AboutShow(); return 0;
                    case IDC_TGL_TRAY: {
                        bool on = !CtrlChecked(m.tglTray);
                        S().trayIcon = on;
                        CtrlSetChecked(m.tglTray, on);
                        SettingsSave();
                        if (on) TrayCreate(); else TrayDestroy();
                        return 0;
                    }
                    case IDC_TGL_AUTOSTART: {
                        bool on = !CtrlChecked(m.tglAuto);
                        std::wstring err;
                        if (AutostartSet(on, err)) {
                            CtrlSetChecked(m.tglAuto, on);
                        } else {
                            CtrlSetChecked(m.tglAuto, !on);   // visual revert on failure
                            MessageBoxW(hwnd, err.c_str(), psqc::Str(STR_APP_TITLE), MB_ICONERROR);
                        }
                        return 0;
                    }
                }
            }
            if (code == STN_CLICKED) {
                if (id == IDC_LBL_TRAY) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_TGL_TRAY, BN_CLICKED), 0); return 0; }
                if (id == IDC_LBL_AUTOSTART) { PostMessageW(hwnd, WM_COMMAND, MAKEWPARAM(IDC_TGL_AUTOSTART, BN_CLICKED), 0); return 0; }
            }
            if (code == CBN_SELCHANGE) {
                if ((HWND)lp == m.cmbLang) {
                    int sel = (int)SendMessageW(m.cmbLang, CB_GETCURSEL, 0, 0);
                    S().lang = sel <= 0 ? std::wstring() : Langs()[sel - 1].tag;
                    SettingsSave();
                    SetLanguage(S().lang.c_str());
                    MainWindowReskin();
                    TrayReskin();
                    if (g.hwndSettings) SettingsReskin();
                    return 0;
                }
                if ((HWND)lp == m.cmbTheme) {
                    int sel = (int)SendMessageW(m.cmbTheme, CB_GETCURSEL, 0, 0);
                    if (sel >= 0 && sel <= 2 && sel != S().themeMode) {
                        S().themeMode = sel;
                        SettingsSave();
                        ThemeSetMode(sel);
                        MainWindowReskin();
                        TrayReskin();
                        if (g.hwndSettings) SettingsReskin();
                    }
                    return 0;
                }
            }
            return 0;
        }
        case WM_NOTIFY: {
            auto* nm = (NMHDR*)lp;
            if (nm->hwndFrom == m.list && nm->code == NM_CUSTOMDRAW) {
                auto* cd = (NMLVCUSTOMDRAW*)lp;
                if (cd->nmcd.dwDrawStage == CDDS_PREPAINT) return CDRF_NOTIFYITEMDRAW;
                if (cd->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
                    cd->clrText = Th().text;
                    cd->clrTextBk = Th().surface2;
                    return CDRF_DODEFAULT;
                }
                return CDRF_DODEFAULT;
            }
            return 0;
        }
        case WM_TIMER:
            if (wp == 1 && !m.busy && IsWindowVisible(m.hwnd)) MainWindowRefreshNow();
            return 0;
        case WM_KEYDOWN: {
            bool ctrl = (GetKeyState(VK_CONTROL) & 0x8000) != 0;
            switch (wp) {
                case VK_F5: MainWindowRefreshNow(); return 0;
                case VK_DELETE: DeleteSelected(); return 0;
                case VK_ESCAPE: PostMessageW(hwnd, WM_CLOSE, 0, 0); return 0;
                case VK_F1: util::OpenUrl(PSQC_WIKI_URL); return 0;
                case 'A': if (ctrl) { ListView_SetItemState(m.list, -1, LVIS_SELECTED, LVIS_SELECTED); return 0; } break;
                case VK_OEM_COMMA: if (ctrl) { SettingsShow(); return 0; } break;
            }
            break;
        }
        case WM_PSQC_REFRESH:
            MainWindowRefreshNow();
            return 0;
        case WM_PSQC_WORK_DONE: {
            auto* res = (spooler::PurgeResult*)lp;
            SetBusy(false);
            MainWindowRefreshNow();
            if (wp == TASK_PURGE_ALL) {
                if (res->ok) {
                    std::wstring msg = util::Fmt(psqc::Str(STR_MSG_CLEAR_OK), util::FormatInt(res->removedBefore).c_str());
                    if (!res->backupPath.empty())
                        msg += L"\n" + util::Fmt(psqc::Str(STR_MSG_BACKUP_SAVED), res->backupPath.c_str());
                    TrayNotify(BK_OK, msg);
                } else {
                    std::wstring msg = util::Fmt(psqc::Str(STR_MSG_CLEAR_FAIL), res->error.c_str());
                    if ((res->errCode == ERROR_ACCESS_DENIED || res->errCode == ERROR_PRIVILEGE_NOT_HELD) &&
                        !g.elevated && S().autoElevate) {
                        if (MessageBoxW(hwnd, psqc::Str(STR_MSG_ELEVATE_Q), psqc::Str(STR_APP_TITLE),
                                        MB_YESNO | MB_ICONWARNING) == IDYES) {
                            DWORD code = 0;
                            util::RunElevatedAndWait(util::ExePath().c_str(), L"/purge", &code);
                            MainWindowRefreshNow();
                        }
                    } else {
                        TrayNotify(BK_ERR, msg);
                    }
                }
            } else if (wp == TASK_RESTART_SPOOLER) {
                if (res->ok) TrayNotify(BK_OK, psqc::Str(STR_MSG_RESTART_OK));
                else TrayNotify(BK_ERR, util::Fmt(psqc::Str(STR_MSG_STOP_FAIL), res->error.c_str()));
            }
            delete res;
            return 0;
        }
        case WM_PSQC_UPDATE_CHECK_DONE: {
            auto* r = (UpdateCheckUi*)lp;
            S().lastUpdateCheck = util::UnixSeconds();
            SettingsSave();
            if (!r->newer) {
                if (r->manual) {
                    if (!r->err.empty())
                        MessageBoxW(hwnd, util::Fmt(psqc::Str(STR_MSG_UPDATE_FAIL), r->err.c_str()).c_str(),
                                    psqc::Str(STR_APP_TITLE), MB_ICONWARNING);
                    else
                        MessageBoxW(hwnd, util::Fmt(psqc::Str(STR_MSG_UPDATE_NONE), PSQC_VERSION_STR_W).c_str(),
                                    psqc::Str(STR_APP_TITLE), MB_ICONINFORMATION);
                }
            } else if (r->info.tag != S().skipVersion) {
                int ans = MessageBoxW(hwnd, util::Fmt(psqc::Str(STR_MSG_UPDATE_AVAIL), r->info.tag.c_str()).c_str(),
                                      psqc::Str(STR_APP_TITLE), MB_YESNO | MB_ICONQUESTION);
                if (ans == IDYES) UpdateApplyAsync(&r->info);
            }
            delete r;
            return 0;
        }
        case WM_PSQC_UPDATE_DL_DONE: {
            auto* r = (UpdateDlUi*)lp;
            if (r->ok) {
                TrayNotify(BK_OK, psqc::Str(STR_UPD_APPLYING));
                if (update::LaunchSelfReplace(r->path, util::ExePath())) {
                    SettingsSave();
                    TrayDestroy();
                    DestroyWindow(hwnd);   // child swaps files and relaunches
                }
            } else {
                MessageBoxW(hwnd, util::Fmt(psqc::Str(STR_MSG_UPDATE_FAIL), r->err.c_str()).c_str(),
                            psqc::Str(STR_APP_TITLE), MB_ICONERROR);
            }
            delete r;
            return 0;
        }
        case WM_PSQC_TRAYICON:
            TrayHandleMessage(wp, lp);
            return 0;
        case WM_PSQC_SINGLETON:
            MainWindowShowAndActivate();
            return 0;
        case WM_CLOSE:
            if (S().trayIcon && S().minimizeToTray) {
                ShowWindow(hwnd, SW_HIDE);
                return 0;
            }
            DestroyWindow(hwnd);
            return 0;
        case WM_DESTROY: {
            RECT rc{};
            GetWindowRect(hwnd, &rc);
            if (!IsIconic(hwnd)) {
                S().winX = rc.left; S().winY = rc.top;
                S().winW = rc.right - rc.left; S().winH = rc.bottom - rc.top;
            }
            S().winMaximized = IsZoomed(hwnd) != 0;
            SettingsSave();
            TrayDestroy();
            m.tips = nullptr;
            PostQuitMessage(0);
            return 0;
        }
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

// ---- create ---------------------------------------------------------------------------
void MainWindowCreate(HINSTANCE hInst, int showCmd) {
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = hInst;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = util::LoadAppIcon(GetSystemMetrics(SM_CXICON), GetSystemMetrics(SM_CYICON));
    wc.hIconSm = util::LoadAppIcon(16, 16);
    wc.lpszClassName = PSQC_WINDOW_CLASS;
    RegisterClassExW(&wc);

    int dpi = util::DpiOf(nullptr);
    int w = MulDiv(1360, dpi, 96), h = MulDiv(620, dpi, 96);   // 8 кнопок тулбара при 10pt требуют ~1300px
    RECT wa = util::WorkAreaRect();
    if (w > (wa.right - wa.left) - 16) w = (wa.right - wa.left) - 16;   // не шире рабочего стола
    int x = S().winX >= 0 ? S().winX : wa.left + ((wa.right - wa.left) - w) / 2;
    int y = S().winY >= 0 ? S().winY : wa.top + ((wa.bottom - wa.top) - h) / 2;
    if (S().winW > 200) w = S().winW;
    if (S().winH > 200) h = S().winH;

    m.hwnd = CreateWindowExW(0, PSQC_WINDOW_CLASS, psqc::Str(STR_APP_TITLE),
                             WS_OVERLAPPEDWINDOW, x, y, w, h, nullptr, nullptr, hInst, nullptr);
    g.hwndMain = m.hwnd;
#ifdef PSQC_QA_DELAY
    Sleep(900);
#endif
    dpi = util::DpiOf(m.hwnd);
    ThemeRefreshDpi(dpi);
    util::SetClientSize(m.hwnd, w, h);   // w x h is the client area, not the frame

    auto mk = [&](int id) { return GetDlgItem(m.hwnd, id); };
    m.list = CreateWindowExW(WS_EX_CLIENTEDGE, WC_LISTVIEWW, L"",
        WS_CHILD | WS_VISIBLE | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SHAREIMAGELISTS,
        0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_LIST, hInst, nullptr);
    ListView_SetExtendedListViewStyle(m.list, LVS_EX_FULLROWSELECT | LVS_EX_DOUBLEBUFFER | LVS_EX_HEADERDRAGDROP);
    if (HWND hdr0 = ListView_GetHeader(m.list)) SetWindowSubclass(hdr0, HeaderProc, 2, 0);
    m.btnRefresh = CreatePsqcButton(m.hwnd, IDC_BTN_REFRESH, STR_TB_REFRESH, GLYPH_REFRESH, BK_Standard);
    m.btnDelSel = CreatePsqcButton(m.hwnd, IDC_BTN_DELSEL, STR_TB_CLEAR_SEL, GLYPH_DELETE, BK_Standard);
    m.btnClearAll = CreatePsqcButton(m.hwnd, IDC_BTN_CLEARALL, STR_TB_CLEAR_ALL, GLYPH_CANCEL, BK_Danger);
    m.btnPause = CreatePsqcButton(m.hwnd, IDC_BTN_PAUSE, STR_TB_PAUSE, GLYPH_PAUSE, BK_Standard);
    m.btnRestart = CreatePsqcButton(m.hwnd, IDC_BTN_RESTART, STR_TB_RESTART, GLYPH_RESTART, BK_Standard);
    m.btnSettings = CreatePsqcButton(m.hwnd, IDC_BTN_SETTINGS, STR_TB_SETTINGS, GLYPH_SETTINGS, BK_Subtle);
    m.btnTheme = CreatePsqcButton(m.hwnd, IDC_BTN_THEME, -1, GLYPH_MOON, BK_Subtle);
    m.btnAbout = CreatePsqcButton(m.hwnd, IDC_BTN_ABOUT, STR_TB_ABOUT, GLYPH_INFO, BK_Subtle);
    CtrlSetText(m.btnTheme, L"");
    m.chipSvc = CreatePsqcChip(m.hwnd, 201);
    m.chipUser = CreatePsqcChip(m.hwnd, 202);
    m.chipPortable = CreatePsqcChip(m.hwnd, 203);
    m.emptyT = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | SS_CENTER, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_EMPTY_TITLE, hInst, nullptr);
    m.emptyH = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | SS_CENTER, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_EMPTY_HINT, hInst, nullptr);
    m.cmbLang = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST | WS_VSCROLL, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_CMB_LANG, hInst, nullptr);
    m.cmbTheme = CreateWindowExW(0, L"COMBOBOX", L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP | CBS_DROPDOWNLIST, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_CMB_THEME, hInst, nullptr);
    m.tglTray = CreatePsqcToggle(m.hwnd, IDC_TGL_TRAY, -1, S().trayIcon);
    m.tglAuto = CreatePsqcToggle(m.hwnd, IDC_TGL_AUTOSTART, -1, AutostartEnabled());
    m.lblLang = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_LBL_LANG, hInst, nullptr);
    m.lblTheme = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_LBL_THEME, hInst, nullptr);
    m.lblTray = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOTIFY, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_LBL_TRAY, hInst, nullptr);
    m.lblAuto = CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOTIFY, 0, 0, 10, 10, m.hwnd, (HMENU)(INT_PTR)IDC_LBL_AUTOSTART, hInst, nullptr);

    ApplyFonts();

    CtrlAddTip(m.btnRefresh, STR_TB_REFRESH);
    CtrlAddTip(m.btnDelSel, STR_TB_CLEAR_SEL);
    CtrlAddTip(m.btnClearAll, STR_TB_CLEAR_ALL);
    CtrlAddTip(m.btnPause, STR_TB_PAUSE);
    CtrlAddTip(m.btnRestart, STR_TB_RESTART);
    CtrlAddTip(m.btnSettings, STR_TB_SETTINGS);
    CtrlAddTip(m.btnTheme, STR_SET_THEME);
    CtrlAddTip(m.btnAbout, STR_TB_ABOUT);

    TipsAdd();
    SetupColumns();
    ApplyStaticTexts();
    MainWindowReskin();
    util::ApplyRtlLayout(m.hwnd);
    Layout();
    MainWindowRefreshNow();

    // theme glyph reflects effective theme
    CtrlSetGlyph(m.btnTheme, Th().dark ? GLYPH_SUN : GLYPH_MOON);

    int show = showCmd;
    if (S().startMinimized || showCmd == SW_MINIMIZE || showCmd == SW_SHOWMINIMIZED) show = SW_HIDE;
    ShowWindow(m.hwnd, show);
    UpdateWindow(m.hwnd);
#ifdef PSQC_QA_DELAY
    Sleep(900);
#endif

    if (S().refreshSec > 0)
        g.timerRefresh = SetTimer(m.hwnd, 1, S().refreshSec * 1000, nullptr);
    UpdateWindow(m.hwnd);
    util::QaShot(m.hwnd, L"main");
}

void MainWindowShowAndActivate() {
    if (!m.hwnd) return;
    ShowWindow(m.hwnd, SW_RESTORE);
    SetForegroundWindow(m.hwnd);
    BringWindowToTop(m.hwnd);
    MainWindowRefreshNow();
}

// ---- update UI glue -----------------------------------------------------------
void UpdateCheckAsync(bool manual) {
    std::thread([manual] {
        auto* r = new UpdateCheckUi();
        r->manual = manual;
        r->newer = update::CheckForUpdate(r->info, r->err);
        PostMessageW(m.hwnd, WM_PSQC_UPDATE_CHECK_DONE, 0, (LPARAM)r);
    }).detach();
}

void UpdateApplyAsync(const update::ReleaseInfo* info) {
    auto inf = *info;
    std::thread([inf] {
        auto* r = new UpdateDlUi();
        r->path = update::UpdateTempPath();
        r->ok = update::DownloadFile(inf.assetUrl, r->path, nullptr, nullptr);
        if (r->ok && util::FileSize(r->path) < 100 * 1024) { r->ok = false; r->err = L"download too small"; }
        if (!r->ok && r->err.empty()) r->err = util::LastErrorText(GetLastError());
        PostMessageW(m.hwnd, WM_PSQC_UPDATE_DL_DONE, 0, (LPARAM)r);
    }).detach();
}
