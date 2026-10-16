// About dialog implementation.
#include "about.h"
#include "i18n.h"
#include "theme.h"
#include "util.h"
#include "settings.h"
#include "controls.h"
#include "main_window.h"

namespace psqc { const wchar_t* Str(int id); }

enum AboutIds : int {
    IDA_TITLE = 301, IDA_VER, IDA_AUTHOR, IDA_COPY,
    IDA_LINK_SITE, IDA_LINK_REPO, IDA_LINK_LIC, IDA_LINK_WIKI,
    IDA_BTN_UPD, IDA_BTN_DIAG, IDA_BTN_CLOSE,
};

struct AB {
    HWND hwnd = nullptr;
    HWND title = nullptr, ver = nullptr, author = nullptr, copy = nullptr;
    HWND lSite = nullptr, lRepo = nullptr, lLic = nullptr, lWiki = nullptr;
    HWND bUpd = nullptr, bDiag = nullptr, bClose = nullptr;
};
static AB a;

static void SetLink(HWND h, int strId, const char* url) {
    std::wstring txt = util::Fmt(L"<a href=\"%S\">%s</a>", url, psqc::Str(strId));
    SetWindowTextW(h, txt.c_str());
}

static void ApplyTexts() {
    SetWindowTextW(a.hwnd, psqc::Str(STR_ABOUT_TITLE));
    SetWindowTextW(a.title, psqc::Str(STR_APP_TITLE));
    SetWindowTextW(a.ver, util::Fmt(psqc::Str(STR_ABOUT_VERSION), PSQC_VERSION_STR_W).c_str());
    SetWindowTextW(a.author, psqc::Str(STR_ABOUT_AUTHOR));
    SetWindowTextW(a.copy, psqc::Str(STR_COPYRIGHT));
    SetLink(a.lSite, STR_LINK_SITE, PSQC_HOMEPAGE);
    SetLink(a.lRepo, STR_LINK_REPO, PSQC_REPO_URL);
    SetLink(a.lLic, STR_LINK_LICENSE, PSQC_LICENSE_URL);
    SetLink(a.lWiki, STR_LINK_WIKI, PSQC_WIKI_URL);
    CtrlSetStrId(a.bUpd, STR_BTN_CHECK_UPD);
    CtrlSetStrId(a.bDiag, STR_BTN_COPY_DIAG);
    CtrlSetStrId(a.bClose, STR_BTN_CLOSE);
}

static std::wstring BuildDiagnostics() {
    return util::Fmt(
        L"%s %s\n%s\nbuild %lu\nelevated: %s\nlang: %s (setting: %s)\ntheme: %d (dark: %s)\nsettings: %s\nportable: %s\n",
        PSQC_PRODUCT_NAME, PSQC_VERSION_STR_W,
        util::OsFriendlyName().c_str(),
        util::WindowsBuild(),
        g.elevated ? L"yes" : L"no",
        Langs()[CurrentLang()].tag, S().lang.empty() ? L"system" : S().lang.c_str(),
        S().themeMode, Th().dark ? L"yes" : L"no",
        util::SettingsFilePath().c_str(),
        g.portable ? L"yes" : L"no");
}

static LRESULT CALLBACK AboutProc(HWND hwnd, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
        case WM_ERASEBKGND: {
            RECT rc{}; GetClientRect(hwnd, &rc);
            FillRect((HDC)wp, &rc, Th().brBg);
            return 1;
        }
        case WM_PAINT: {
            PAINTSTRUCT ps{};
            HDC hdc = BeginPaint(hwnd, &ps);
            int dpi = util::DpiOf(hwnd);
            HICON ico = util::LoadAppIcon(MulDiv(48, dpi, 96), MulDiv(48, dpi, 96));
            DrawIconEx(hdc, MulDiv(20, dpi, 96), MulDiv(20, dpi, 96), ico, MulDiv(48, dpi, 96), MulDiv(48, dpi, 96), 0, nullptr, DI_NORMAL);
            EndPaint(hwnd, &ps);
            return 0;
        }
        case WM_CTLCOLORSTATIC:
            if ((HWND)lp == a.lSite || (HWND)lp == a.lRepo || (HWND)lp == a.lLic || (HWND)lp == a.lWiki) {
                SetBkColor((HDC)wp, Th().bg);
                SetTextColor((HDC)wp, Th().accent);
                return (LRESULT)Th().brBg;
            }
            return (LRESULT)ThemeCtlColorBg((HDC)wp);
        case WM_DRAWITEM:
            if (CtrlHandleDrawItem((DRAWITEMSTRUCT*)lp)) return TRUE;
            return 0;
        case WM_COMMAND: {
            int id = LOWORD(wp);
            if (HIWORD(wp) == BN_CLICKED) {
                if (id == IDA_BTN_CLOSE) { DestroyWindow(hwnd); return 0; }
                if (id == IDA_BTN_UPD) { UpdateCheckAsync(true); return 0; }
                if (id == IDA_BTN_DIAG) {
                    util::CopyToClipboard(BuildDiagnostics());
                    MessageBoxW(hwnd, psqc::Str(STR_MSG_DIAG_COPIED), psqc::Str(STR_APP_TITLE), MB_ICONINFORMATION);
                    return 0;
                }
            }
            return 0;
        }
        case WM_NOTIFY: {
            auto* nm = (NMHDR*)lp;
            if (nm->code == NM_CLICK || nm->code == NM_RETURN) {
                auto* li = (NMLINK*)lp;
                if (li->item.szUrl[0]) util::OpenUrl(util::WideToUtf8(li->item.szUrl).c_str());
            }
            return 0;
        }
        case WM_PSQC_LANG_CHANGED:
            ApplyTexts();
            InvalidateRect(hwnd, nullptr, TRUE);
            return 0;
        case WM_CLOSE:
            DestroyWindow(hwnd);
            return 0;
        case WM_DESTROY:
            g.hwndAbout = nullptr;
            return 0;
    }
    return DefWindowProcW(hwnd, msg, wp, lp);
}

void AboutShow() {
    if (g.hwndAbout) { SetForegroundWindow(g.hwndAbout); return; }
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    wc.lpfnWndProc = AboutProc;
    wc.hInstance = g.hInst;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hIcon = util::LoadAppIcon(32, 32);
    wc.lpszClassName = PSQC_ABOUT_CLASS;
    RegisterClassExW(&wc);

    int dpi = util::DpiOf(g.hwndMain);
    auto S = [&](int v) { return MulDiv(v, dpi, 96); };
    int W = S(470), H = S(430);
    a.hwnd = CreateWindowExW(WS_EX_DLGMODALFRAME, PSQC_ABOUT_CLASS, psqc::Str(STR_ABOUT_TITLE),
                             WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU, 0, 0, W, H,
                             g.hwndMain, nullptr, g.hInst, nullptr);
    g.hwndAbout = a.hwnd;
    ThemeDarkTitlebar(a.hwnd, Th().dark);

    auto mkStatic = [&](int id, DWORD style) {
        return CreateWindowExW(0, L"STATIC", L"", WS_CHILD | WS_VISIBLE | style, 0, 0, 10, 10, a.hwnd, (HMENU)(INT_PTR)id, g.hInst, nullptr);
    };
    a.title = mkStatic(IDA_TITLE, SS_LEFT);
    a.ver = mkStatic(IDA_VER, SS_LEFT);
    a.author = mkStatic(IDA_AUTHOR, SS_LEFT);
    a.copy = mkStatic(IDA_COPY, SS_LEFT);
    auto mkLink = [&](int id) {
        return CreateWindowExW(0, WC_LINK, L"", WS_CHILD | WS_VISIBLE | WS_TABSTOP, 0, 0, 10, 10, a.hwnd, (HMENU)(INT_PTR)id, g.hInst, nullptr);
    };
    a.lSite = mkLink(IDA_LINK_SITE);
    a.lRepo = mkLink(IDA_LINK_REPO);
    a.lLic = mkLink(IDA_LINK_LIC);
    a.lWiki = mkLink(IDA_LINK_WIKI);
    a.bUpd = CreatePsqcButton(a.hwnd, IDA_BTN_UPD, STR_BTN_CHECK_UPD, GLYPH_DOWNLOAD, BK_Standard);
    a.bDiag = CreatePsqcButton(a.hwnd, IDA_BTN_DIAG, STR_BTN_COPY_DIAG, GLYPH_COPY, BK_Standard);
    a.bClose = CreatePsqcButton(a.hwnd, IDA_BTN_CLOSE, STR_BTN_CLOSE, 0, BK_Accent);

    SendMessageW(a.title, WM_SETFONT, (WPARAM)Th().fontBig, TRUE);
    SendMessageW(a.ver, WM_SETFONT, (WPARAM)Th().font, TRUE);
    SendMessageW(a.author, WM_SETFONT, (WPARAM)Th().font, TRUE);
    SendMessageW(a.copy, WM_SETFONT, (WPARAM)Th().fontSmall, TRUE);
    for (HWND h : { a.lSite, a.lRepo, a.lLic, a.lWiki }) SendMessageW(h, WM_SETFONT, (WPARAM)Th().font, TRUE);

    ApplyTexts();
    ThemeDarkControls(a.hwnd, Th().dark);
    util::ApplyRtlLayout(a.hwnd);

    // layout
    int x = S(80), y = S(24);
    SetWindowPos(a.title, nullptr, x, y, W - x - S(20), S(28), SWP_NOZORDER); y += S(30);
    SetWindowPos(a.ver, nullptr, x, y, W - x - S(20), S(20), SWP_NOZORDER); y += S(24);
    SetWindowPos(a.author, nullptr, x, y, W - x - S(20), S(20), SWP_NOZORDER); y += S(20);
    SetWindowPos(a.copy, nullptr, x, y, W - x - S(20), S(18), SWP_NOZORDER); y += S(28);
    for (HWND h : { a.lSite, a.lRepo, a.lLic, a.lWiki }) {
        SetWindowPos(h, nullptr, x, y, W - x - S(20), S(20), SWP_NOZORDER);
        y += S(22);
    }
    y += S(8);
    int bw = S(150), bh = S(32);
    SetWindowPos(a.bUpd, nullptr, S(20), y, bw, bh, SWP_NOZORDER);
    SetWindowPos(a.bDiag, nullptr, S(20) + bw + S(8), y, bw, bh, SWP_NOZORDER);
    SetWindowPos(a.bClose, nullptr, W - S(20) - S(100), y, S(100), bh, SWP_NOZORDER);

    util::CenterWindow(a.hwnd, g.hwndMain);
    ShowWindow(a.hwnd, SW_SHOW);
}
