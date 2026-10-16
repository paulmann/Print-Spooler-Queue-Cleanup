// Custom controls implementation.
#include "controls.h"
#include "theme.h"
#include "i18n.h"
#include "util.h"
#include <commctrl.h>

extern HINSTANCE g_hInstForControls();   // defined in main.cpp (returns g.hInst)

struct CtrlData {
    int kind = CK_Button;
    int btnKind = BK_Standard;
    int strId = -1;
    wchar_t glyph = 0;
    bool checked = false;
    bool hot = false;
    bool pressed = false;
    bool focused = false;
    bool hasTextOverride = false;
    std::wstring textOverride;
    COLORREF dot = 0;
    int tipStrId = -1;
    HWND tip = nullptr;
};

static std::map<HWND, CtrlData> g_ctrls;

static CtrlData* Data(HWND h) {
    auto it = g_ctrls.find(h);
    return it == g_ctrls.end() ? nullptr : &it->second;
}

static std::wstring LabelOf(const CtrlData& d) {
    if (d.hasTextOverride) return d.textOverride;
    if (d.strId >= 0) return psqc::Str(d.strId);
    return {};
}

static LRESULT CALLBACK CtrlSubclass(HWND h, UINT msg, WPARAM wp, LPARAM lp, UINT_PTR, DWORD_PTR) {
    CtrlData* d = Data(h);
    switch (msg) {
        case WM_MOUSEMOVE:
            if (d && !d->hot) {
                d->hot = true;
                TRACKMOUSEEVENT tme{ sizeof(tme), TME_LEAVE, h, 0 };
                TrackMouseEvent(&tme);
                InvalidateRect(h, nullptr, FALSE);
            }
            break;
        case WM_MOUSELEAVE:
            if (d && d->hot) { d->hot = false; InvalidateRect(h, nullptr, FALSE); }
            break;
        case WM_LBUTTONDOWN:
            if (d) { d->pressed = true; SetCapture(h); InvalidateRect(h, nullptr, FALSE); }
            break;
        case WM_LBUTTONUP:
            if (d && d->pressed) {
                d->pressed = false;
                ReleaseCapture();
                InvalidateRect(h, nullptr, FALSE);
            }
            break;
        case WM_SETFOCUS:
            if (d) { d->focused = true; InvalidateRect(h, nullptr, FALSE); }
            break;
        case WM_KILLFOCUS:
            if (d) { d->focused = false; InvalidateRect(h, nullptr, FALSE); }
            break;
        case WM_NCDESTROY:
            RemoveWindowSubclass(h, CtrlSubclass, 1);
            g_ctrls.erase(h);
            break;
    }
    return DefSubclassProc(h, msg, wp, lp);
}

static HWND MakeBase(HWND parent, int ctrlId, DWORD style, int w, int hh) {
#ifdef PSQC_QA_DELAY
    Sleep(150);   // QA-only build knob: slows creation down for Wine race workaround
#endif
    // NOTE: no WS_VISIBLE here — the first WM_DRAWITEM must happen only after the
    // control's paint data is registered (otherwise the first paint draws nothing and,
    // on some comctl implementations, an unhandled owner-draw item is never repainted).
    HWND h = CreateWindowExW(0, L"BUTTON", L"", style | WS_CHILD | BS_OWNERDRAW,
                             0, 0, w, hh, parent, (HMENU)(INT_PTR)ctrlId, g_hInstForControls(), nullptr);
    SetWindowSubclass(h, CtrlSubclass, 1, 0);
    return h;
}

// hInst plumbed via app state to avoid include cycle
extern HINSTANCE g_hInstForControls();

HWND CreatePsqcButton(HWND parent, int ctrlId, int strId, wchar_t glyph, int btnKind) {
    HWND h = MakeBase(parent, ctrlId, WS_TABSTOP, 120, 32);
    CtrlData d;
    d.kind = CK_Button;
    d.btnKind = btnKind;
    d.strId = strId;
    d.glyph = glyph;
    g_ctrls[h] = d;
    ShowWindow(h, SW_SHOW);   // only now the first WM_DRAWITEM can paint correctly
    return h;
}

HWND CreatePsqcToggle(HWND parent, int ctrlId, int strId, bool checked) {
    HWND h = MakeBase(parent, ctrlId, WS_TABSTOP, 46, 22);
    CtrlData d;
    d.kind = CK_Toggle;
    d.strId = strId;
    d.checked = checked;
    g_ctrls[h] = d;
    ShowWindow(h, SW_SHOW);
    return h;
}

HWND CreatePsqcChip(HWND parent, int ctrlId) {
    HWND h = MakeBase(parent, ctrlId, 0, 90, 24);
    CtrlData d;
    d.kind = CK_Chip;
    g_ctrls[h] = d;
    ShowWindow(h, SW_SHOW);
    return h;
}

void CtrlSetChecked(HWND h, bool on) { if (auto* d = Data(h)) { d->checked = on; InvalidateRect(h, nullptr, TRUE); } }
bool CtrlChecked(HWND h) { auto* d = Data(h); return d && d->checked; }
void CtrlSetStrId(HWND h, int strId) { if (auto* d = Data(h)) { d->strId = strId; d->hasTextOverride = false; InvalidateRect(h, nullptr, TRUE); } }
void CtrlSetText(HWND h, const std::wstring& text) { if (auto* d = Data(h)) { d->textOverride = text; d->hasTextOverride = true; InvalidateRect(h, nullptr, TRUE); } }
std::wstring CtrlText(HWND h) { auto* d = Data(h); return d ? LabelOf(*d) : std::wstring(); }
void CtrlSetGlyph(HWND h, wchar_t glyph) { if (auto* d = Data(h)) { d->glyph = glyph; InvalidateRect(h, nullptr, TRUE); } }
void CtrlSetKind(HWND h, int btnKind) { if (auto* d = Data(h)) { d->btnKind = btnKind; InvalidateRect(h, nullptr, TRUE); } }
void CtrlSetDotColor(HWND h, COLORREF c) { if (auto* d = Data(h)) { d->dot = c; InvalidateRect(h, nullptr, TRUE); } }
void CtrlSetEnabledVisual(HWND h, bool en) { EnableWindow(h, en); InvalidateRect(h, nullptr, TRUE); }

void CtrlRelabelAll() {
    for (auto& kv : g_ctrls) InvalidateRect(kv.first, nullptr, TRUE);
}

void CtrlAddTip(HWND h, int strId) {
    auto* d = Data(h);
    if (!d) return;
    d->tipStrId = strId;
    HWND parent = GetParent(h);
    if (!d->tip && parent) {
        d->tip = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr,
                                 WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP | TTF_SUBCLASS,
                                 0, 0, 0, 0, parent, nullptr, g_hInstForControls(), nullptr);
        if (d->tip) {
            SetWindowPos(d->tip, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
            TOOLINFOW ti{};
            ti.cbSize = sizeof(ti);
            ti.uFlags = TTF_SUBCLASS | TTF_IDISHWND;
            ti.hwnd = parent;
            ti.uId = (UINT_PTR)h;
            ti.lpszText = (LPWSTR)psqc::Str(strId);
            SendMessageW(d->tip, TTM_ADDTOOLW, 0, (LPARAM)&ti);
            SendMessageW(d->tip, TTM_SETMAXTIPWIDTH, 0, 400);
        }
    } else if (d->tip) {
        TOOLINFOW ti{};
        ti.cbSize = sizeof(ti);
        ti.uFlags = TTF_SUBCLASS | TTF_IDISHWND;
        ti.hwnd = GetParent(h);
        ti.uId = (UINT_PTR)h;
        ti.lpszText = (LPWSTR)psqc::Str(strId);
        SendMessageW(d->tip, TTM_UPDATETIPTEXTW, 0, (LPARAM)&ti);
    }
}

void CtrlResizeToContent(HWND h, HDC hdcOpt) {
    auto* d = Data(h);
    if (!d) return;
    HDC hdc = hdcOpt ? hdcOpt : GetDC(h);
    HFONT old = (HFONT)SelectObject(hdc, Th().font);
    auto label = LabelOf(*d);
    SIZE ts{};
    GetTextExtentPoint32W(hdc, label.c_str(), (int)label.size(), &ts);
    SelectObject(hdc, old);
    if (!hdcOpt) ReleaseDC(h, hdc);
    int pad = util::Scale(h, 14);
    int gh = d->glyph ? util::Scale(h, 18) : 0;
    int w = pad * 2 + gh + (d->glyph && !label.empty() ? util::Scale(h, 6) : 0) + ts.cx;
    int hh = d->kind == CK_Toggle ? util::Scale(h, 22) : d->kind == CK_Chip ? util::Scale(h, 24) : util::Scale(h, 32);
    if (d->kind == CK_Toggle) w = util::Scale(h, 44);
    SetWindowPos(h, nullptr, 0, 0, w, hh, SWP_NOMOVE | SWP_NOZORDER);
}

// ---- painting ----------------------------------------------------------------
static void DrawRoundRect(HDC hdc, const RECT& r, int rad) {
    BeginPath(hdc);
    RoundRect(hdc, r.left, r.top, r.right + 1, r.bottom + 1, rad, rad);
    EndPath(hdc);
}

static void FillRound(HDC hdc, const RECT& r, int rad, COLORREF fill, COLORREF border, bool hasBorder) {
    HBRUSH br = CreateSolidBrush(fill);
    HPEN pen = hasBorder ? CreatePen(PS_SOLID, 1, border) : (HPEN)GetStockObject(NULL_PEN);
    HGDIOBJ ob = SelectObject(hdc, br);
    HGDIOBJ op = SelectObject(hdc, pen);
    DrawRoundRect(hdc, r, rad);
    FillPath(hdc);
    SelectObject(hdc, ob);
    SelectObject(hdc, op);
    DeleteObject(br);
    if (hasBorder) DeleteObject(pen);
}

static COLORREF Mix(COLORREF a, COLORREF b, double t) {
    auto R = [](COLORREF c, int sh) { return (c >> sh) & 0xFF; };
    return RGB((BYTE)(R(a, 0) + (R(b, 0) - R(a, 0)) * t),
               (BYTE)(R(a, 8) + (R(b, 8) - R(a, 8)) * t),
               (BYTE)(R(a, 16) + (R(b, 16) - R(a, 16)) * t));
}

bool CtrlHandleDrawItem(DRAWITEMSTRUCT* dis) {
    auto* d = Data(dis->hwndItem);
    if (!d) return false;
    const Theme& t = Th();
    HDC hdc = dis->hDC;
    RECT r = dis->rcItem;
    bool enabled = (dis->itemState & ODS_DISABLED) == 0;
    int rad = util::Scale(dis->hwndItem, d->kind == CK_Chip ? 12 : 6);

    if (d->kind == CK_Toggle) {
        // track + knob, label drawn by parent static; switch fills the control
        int hh = r.bottom - r.top;
        int tw = hh * 2;
        RECT tr{ r.left, r.top, r.left + tw, r.bottom };
        COLORREF track = d->checked ? t.accent : t.surface2;
        COLORREF bord = d->checked ? t.accent : t.border;
        FillRound(hdc, tr, hh / 2, track, bord, true);
        int knob = hh - util::Scale(dis->hwndItem, 8);
        int kx = d->checked ? tr.right - knob - util::Scale(dis->hwndItem, 4)
                            : tr.left + util::Scale(dis->hwndItem, 4);
        RECT kr{ kx, r.top + util::Scale(dis->hwndItem, 4), kx + knob, r.bottom - util::Scale(dis->hwndItem, 4) };
        COLORREF kc = d->checked ? t.accentText : t.subtext;
        if (d->hot && enabled) kc = d->checked ? Mix(kc, t.accentText, 0) : Mix(kc, t.text, 0.3);
        FillRound(hdc, kr, knob / 2, kc, kc, false);
        if (d->focused) { DrawFocusRect(hdc, &r); }
        return true;
    }

    if (d->kind == CK_Chip) {
        FillRound(hdc, r, rad, t.chipBg, t.chipBg, false);
        SetBkMode(hdc, TRANSPARENT);
        auto label = LabelOf(*d);
        HFONT of = (HFONT)SelectObject(hdc, t.fontSmall);
        SetTextColor(hdc, t.text);
        int dotX = r.left + util::Scale(dis->hwndItem, 8);
        if (d->dot) {
            int cy = (r.top + r.bottom) / 2;
            RECT dr{ dotX, cy - 4, dotX + 8, cy + 4 };
            FillRound(hdc, dr, 4, d->dot, d->dot, false);
            dotX += 12;
        }
        RECT tr2{ dotX, r.top, r.right - 6, r.bottom };
        DrawTextW(hdc, label.c_str(), -1, &tr2, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
        SelectObject(hdc, of);
        return true;
    }

    // button
    COLORREF fill = t.surface2, border = t.border, fg = t.text;
    switch (d->btnKind) {
        case BK_Accent: fill = t.accent; border = t.accent; fg = t.accentText; break;
        case BK_Danger: fill = t.danger; border = t.danger; fg = RGB(255, 255, 255); break;
        case BK_Subtle: fill = t.bg; border = t.bg; fg = t.text; break;
        default: break;
    }
    if (enabled && d->hot) {
        if (d->btnKind == BK_Accent) fill = t.accentHover;
        else if (d->btnKind == BK_Danger) fill = t.dangerHover;
        else if (d->btnKind == BK_Subtle) fill = t.hoverBg;
        else fill = t.hoverBg;
    }
    if (enabled && d->pressed) fill = Mix(fill, t.bg, 0.25);
    if (!enabled) { fill = Mix(fill, t.bg, 0.5); fg = t.subtext; border = Mix(border, t.bg, 0.6); }
    FillRound(hdc, r, rad, fill, border, d->btnKind == BK_Standard || d->btnKind == BK_Subtle);

    auto label = LabelOf(*d);
    SetBkMode(hdc, TRANSPARENT);
    SetTextColor(hdc, fg);
    // measure to center glyph+label
    HFONT of = (HFONT)SelectObject(hdc, t.font);
    SIZE ts{};
    GetTextExtentPoint32W(hdc, label.c_str(), (int)label.size(), &ts);
    int gh = d->glyph ? util::Scale(dis->hwndItem, 16) : 0;
    int total = gh + (d->glyph && !label.empty() ? util::Scale(dis->hwndItem, 6) : 0) + ts.cx;
    int x = r.left + ((r.right - r.left) - total) / 2;
    if (d->glyph) {
        HFONT og = (HFONT)SelectObject(hdc, t.fontIcon);
        wchar_t gbuf[2] = { d->glyph, 0 };
        RECT gr{ x, r.top, x + gh, r.bottom };
        DrawTextW(hdc, gbuf, 1, &gr, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
        SelectObject(hdc, og);
        x += gh + util::Scale(dis->hwndItem, 6);
    }
    if (!label.empty()) {
        RECT tr3{ x, r.top, x + ts.cx, r.bottom };
        DrawTextW(hdc, label.c_str(), -1, &tr3, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    }
    SelectObject(hdc, of);
    if (d->focused && (dis->itemState & ODS_FOCUS)) {
        RECT fr = r;
        InflateRect(&fr, -3, -3);
        DrawFocusRect(hdc, &fr);
    }
    return true;
}
