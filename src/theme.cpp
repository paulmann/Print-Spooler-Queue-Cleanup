// Theming implementation (Fluent-ish palettes, Win10/11 immersive dark mode).
#include "theme.h"
#include "util.h"
#include "settings.h"

static Theme g_theme;
static int   g_mode = THEME_SYSTEM;
static int   g_dpi = 96;

typedef BOOL(WINAPI* PFN_AllowDarkModeForWindow)(HWND, BOOL);
typedef int(WINAPI* PFN_SetPreferredAppMode)(int);
typedef void(WINAPI* PFN_FlushMenuThemes)();
typedef void(WINAPI* PFN_RefreshImmersiveColorPolicyState)();
static PFN_AllowDarkModeForWindow pAllowDark = nullptr;
static PFN_SetPreferredAppMode pAppMode = nullptr;
static PFN_FlushMenuThemes pFlushMenus = nullptr;
static PFN_RefreshImmersiveColorPolicyState pRefreshPolicy = nullptr;

static void LoadUxtheme() {
    static bool done = false;
    if (done) return;
    done = true;
    // The immersive-dark ordinals (104/133/135/136) are undocumented and their meaning
    // differs across Windows builds and Wine implementations. On Wine they map to
    // unrelated functions and calling them corrupts the theme manager, so skip them
    // there (documented SetWindowTheme/DWM paths still give full dark mode).
    if (util::IsWineRuntime()) return;
    HMODULE ux = LoadLibraryW(L"uxtheme.dll");
    if (!ux) return;
    pAllowDark = (PFN_AllowDarkModeForWindow)GetProcAddress(ux, MAKEINTRESOURCEA(133));
    pAppMode = (PFN_SetPreferredAppMode)GetProcAddress(ux, MAKEINTRESOURCEA(135));
    pFlushMenus = (PFN_FlushMenuThemes)GetProcAddress(ux, MAKEINTRESOURCEA(136));
    pRefreshPolicy = (PFN_RefreshImmersiveColorPolicyState)GetProcAddress(ux, MAKEINTRESOURCEA(104));
    if (pAppMode) pAppMode(1);   // AllowDark
    if (pRefreshPolicy) pRefreshPolicy();
}

bool SystemIsDark() {
    DWORD v = 1;
    if (util::RegGetDword(HKEY_CURRENT_USER, L"Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize",
                          L"AppsUseLightTheme", v))
        return v == 0;
    return false;
}

bool ThemeIsDarkEffective() {
    if (g_mode == THEME_SYSTEM) return SystemIsDark();
    return g_mode == THEME_DARK;
}

int ThemeModeNow() { return g_mode; }

// Icon fonts: Windows 11 moved the full icon set to "Segoe Fluent Icons" and keeps
// "Segoe MDL2 Assets" as a reduced compatibility font (measured on build 26200: U+E8B7
// "Folder" is an empty box there, while Fluent Icons draws the folder; every other glyph
// this app uses is identical in both). Prefer Fluent, fall back to MDL2 on Windows 10.
static bool FontFaceInstalled(const wchar_t* face) {
    HDC dc = GetDC(nullptr);
    HFONT f = CreateFontW(-12, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                          OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                          DEFAULT_PITCH | FF_SWISS, face);
    HGDIOBJ old = SelectObject(dc, f);
    wchar_t actual[LF_FACESIZE]{};
    GetTextFaceW(dc, LF_FACESIZE, actual);   // a missing face is silently substituted
    SelectObject(dc, old);
    DeleteObject(f);
    ReleaseDC(nullptr, dc);
    return _wcsicmp(actual, face) == 0;
}

static const wchar_t* IconFace() {
    static const wchar_t* cached = nullptr;
    if (!cached)
        cached = FontFaceInstalled(L"Segoe Fluent Icons") ? L"Segoe Fluent Icons" : L"Segoe MDL2 Assets";
    return cached;
}

static HFONT MakeFont(int dpi, int pt, bool bold, const wchar_t* face) {
    // Segoe UI Variable exists on Windows 11 only; fall back to Segoe UI on Windows 10.
    if (wcscmp(face, L"Segoe UI Variable Text") == 0 || wcscmp(face, L"Segoe UI Variable Display") == 0) {
        DWORD build = 0;
        // CurrentBuildNumber — REG_SZ: числовой буфер 4 байта всегда давал build=0
        wchar_t num[32]{};
        DWORD sz = sizeof(num), type = 0;
        HKEY k{};
        if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", 0, KEY_READ, &k) == ERROR_SUCCESS) {
            if (RegQueryValueExW(k, L"CurrentBuildNumber", nullptr, &type, (BYTE*)num, &sz) == ERROR_SUCCESS && type == REG_SZ)
                build = (DWORD)wcstoul(num, nullptr, 10);
            RegCloseKey(k);
        }
        if (build && build < 22000)
            face = (wcscmp(face, L"Segoe UI Variable Display") == 0) ? L"Segoe UI Semibold" : L"Segoe UI";
    }
    HFONT f = CreateFontW(-MulDiv(pt, dpi, 72), 0, 0, 0, bold ? FW_SEMIBOLD : FW_NORMAL, FALSE, FALSE, FALSE,
                          DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                          DEFAULT_PITCH | FF_SWISS, face);
    // Запрошенного семейства нет, и GDI молча подставил другое — тогда вместо
    // кириллицы в некоторых шрифтах рисуются вертикальные палки. Проверяем
    // фактическое начертание и при подмене возвращаемся к Segoe UI.
    HDC dc = GetDC(nullptr);
    HGDIOBJ old = SelectObject(dc, f);
    wchar_t actual[LF_FACESIZE]{};
    GetTextFaceW(dc, LF_FACESIZE, actual);
    SelectObject(dc, old);
    ReleaseDC(nullptr, dc);
    if (_wcsicmp(actual, face) != 0) {
        DeleteObject(f);
        f = CreateFontW(-MulDiv(pt, dpi, 72), 0, 0, 0, bold ? FW_SEMIBOLD : FW_NORMAL, FALSE, FALSE, FALSE,
                        DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
                        DEFAULT_PITCH | FF_SWISS, L"Segoe UI");
    }
    return f;
}

static void FreeTheme(Theme& t) {
    for (HBRUSH* b : { &t.brBg, &t.brSurface, &t.brSurface2, &t.brHeader, &t.brBorder }) {
        if (*b) DeleteObject(*b);
        *b = nullptr;
    }
    for (HFONT* f : { &t.font, &t.fontBold, &t.fontSmall, &t.fontBig, &t.fontIcon, &t.fontIconBig }) {
        if (*f) DeleteObject(*f);
        *f = nullptr;
    }
}

static void BuildPalette(Theme& t, bool dark, int dpi) {
    FreeTheme(t);
    t.dark = dark;
    // DeepSeek-подобная палитра: спокойный графит, синий акцент #4D6BFE,
    // читаемые рамки (в тёмной теме рамки светлее фона, а не темнее).
    if (dark) {
        t.bg = RGB(0x1F, 0x20, 0x23);
        t.surface = RGB(0x26, 0x27, 0x2B);
        t.surface2 = RGB(0x2A, 0x2C, 0x31);
        t.text = RGB(0xEC, 0xED, 0xF0);
        t.subtext = RGB(0xA5, 0xA8, 0xAE);
        t.border = RGB(0x45, 0x48, 0x4F);
        t.accent = RGB(0x4D, 0x6B, 0xFE);
        t.accentHover = RGB(0x6C, 0x85, 0xFF);
        t.accentText = RGB(0xFF, 0xFF, 0xFF);
        t.danger = RGB(0xE5, 0x48, 0x4D);
        t.dangerHover = RGB(0xF0, 0x5B, 0x60);
        t.ok = RGB(0x4C, 0xC3, 0x8A);
        t.warn = RGB(0xF5, 0xA5, 0x24);
        t.err = RGB(0xFF, 0x63, 0x69);
        t.headerBg = RGB(0x38, 0x3B, 0x42);
        t.selBg = RGB(0x33, 0x41, 0x5E);
        t.hoverBg = RGB(0x33, 0x36, 0x3C);
        t.chipBg = RGB(0x38, 0x3B, 0x42);
    } else {
        t.bg = RGB(0xF7, 0xF8, 0xFA);
        t.surface = RGB(0xFF, 0xFF, 0xFF);
        t.surface2 = RGB(0xFF, 0xFF, 0xFF);
        t.text = RGB(0x26, 0x28, 0x2C);
        t.subtext = RGB(0x6E, 0x72, 0x78);
        t.border = RGB(0xE1, 0xE3, 0xE8);
        t.accent = RGB(0x4D, 0x6B, 0xFE);
        t.accentHover = RGB(0x3D, 0x5A, 0xF1);
        t.accentText = RGB(0xFF, 0xFF, 0xFF);
        t.danger = RGB(0xE5, 0x48, 0x4D);
        t.dangerHover = RGB(0xC9, 0x3A, 0x3F);
        t.ok = RGB(0x17, 0x93, 0x4B);
        t.warn = RGB(0xB4, 0x53, 0x09);
        t.err = RGB(0xE5, 0x48, 0x4D);
        t.headerBg = RGB(0xE8, 0xEA, 0xEE);
        t.selBg = RGB(0xE3, 0xEB, 0xFF);
        t.hoverBg = RGB(0xEE, 0xF0, 0xF4);
        t.chipBg = RGB(0xEC, 0xEE, 0xF2);
    }
    t.brBg = CreateSolidBrush(t.bg);
    t.brSurface = CreateSolidBrush(t.surface);
    t.brSurface2 = CreateSolidBrush(t.surface2);
    t.brHeader = CreateSolidBrush(t.headerBg);
    t.brBorder = CreateSolidBrush(t.border);
    // Шрифт интерфейса — настройка пользователя; по умолчанию Segoe UI: статический,
    // полная кириллица/греческая азбука + системный font-linking для CJK/арабицы/иврита.
    const wchar_t* face = S().fontFace.empty() ? L"Segoe UI" : S().fontFace.c_str();
    int fsz = (S().fontSize >= 8 && S().fontSize <= 16) ? S().fontSize : 10;
    int ssz = fsz - 1 < 8 ? 8 : fsz - 1;
    t.font = MakeFont(dpi, fsz, false, face);
    t.fontBold = MakeFont(dpi, fsz, true, face);
    t.fontSmall = MakeFont(dpi, ssz, false, face);
    t.fontBig = MakeFont(dpi, fsz + 4, true, face);
    t.fontIcon = MakeFont(dpi, 10, false, IconFace());
    t.fontIconBig = MakeFont(dpi, 16, false, IconFace());
}

const Theme& Th() { return g_theme; }
bool ThemeDarkNow() { return g_theme.dark; }

void ThemeInit(int dpi) {
    LoadUxtheme();
    g_mode = S().themeMode;   // стартуем в сохранённом режиме (раньше настройка применялась только после первого клика)
    g_dpi = dpi;
    BuildPalette(g_theme, ThemeIsDarkEffective(), dpi);
}

void ThemeRefreshDpi(int dpi) {
    if (dpi == g_dpi) return;
    g_dpi = dpi;
    BuildPalette(g_theme, g_theme.dark, dpi);
}

void ThemeSetMode(int mode) {
    g_mode = mode;
    BuildPalette(g_theme, ThemeIsDarkEffective(), g_dpi);
    if (pRefreshPolicy) pRefreshPolicy();
    if (pFlushMenus) pFlushMenus();
}

void ThemeDarkTitlebar(HWND hwnd, bool dark) {
    BOOL v = dark ? TRUE : FALSE;
    ::DwmSetWindowAttribute(hwnd, 20 /*DWMWA_USE_IMMERSIVE_DARK_MODE*/, &v, sizeof(v));
    if (pAllowDark) pAllowDark(hwnd, v);
}

static BOOL CALLBACK DarkenChild(HWND child, LPARAM lp) {
    bool dark = (lp != 0);
    wchar_t cls[64]{};
    GetClassNameW(child, cls, 64);
    const wchar_t* theme = dark ? L"DarkMode_Explorer" : L"Explorer";
    if (wcscmp(cls, WC_LISTVIEWW) == 0 || wcscmp(cls, L"ComboBox") == 0 || wcscmp(cls, L"Edit") == 0) {
        SetWindowTheme(child, theme, nullptr);
    } else if (wcscmp(cls, L"SysHeader32") == 0) {
        SetWindowTheme(child, dark ? L"DarkMode_ItemsView" : L"ItemsView", nullptr);
    } else if (wcscmp(cls, L"SysLink") == 0) {
        SetWindowTheme(child, dark ? L"DarkMode_Explorer" : nullptr, nullptr);
    }
    InvalidateRect(child, nullptr, TRUE);
    return TRUE;
}

void ThemeDarkControls(HWND hwnd, bool dark) {
    EnumChildWindows(hwnd, DarkenChild, dark ? 1 : 0);
}

HBRUSH ThemeCtlColorBg(HDC hdc) {
    SetBkColor(hdc, g_theme.bg);
    SetTextColor(hdc, g_theme.text);
    return g_theme.brBg;
}
