#pragma once
// Theming: light/dark palettes, fonts, system-theme auto-detection, dark-mode plumbing.
#include "app.h"

struct Theme {
    bool dark = false;
    COLORREF bg          = 0;   // window background
    COLORREF surface     = 0;   // cards / panels
    COLORREF surface2    = 0;   // inputs, list background
    COLORREF text        = 0;
    COLORREF subtext     = 0;
    COLORREF border      = 0;
    COLORREF accent      = 0;   // primary fill
    COLORREF accentHover = 0;
    COLORREF accentText  = 0;   // text on accent
    COLORREF danger      = 0;
    COLORREF dangerHover = 0;
    COLORREF ok          = 0;
    COLORREF warn        = 0;
    COLORREF err         = 0;
    COLORREF headerBg    = 0;   // list header
    COLORREF selBg       = 0;   // list selection
    COLORREF hoverBg     = 0;
    COLORREF chipBg      = 0;
    HBRUSH brBg = nullptr, brSurface = nullptr, brSurface2 = nullptr;
    HFONT font = nullptr, fontBold = nullptr, fontSmall = nullptr, fontBig = nullptr;
    HFONT fontIcon = nullptr, fontIconBig = nullptr;   // Segoe MDL2 Assets
};

const Theme& Th();
void ThemeInit(int dpi);                 // (re)create palettes+fonts for dpi & current mode
bool ThemeIsDarkEffective();             // resolved: mode==system ? registry : mode
bool SystemIsDark();                     // HKCU AppsUseLightTheme == 0
void ThemeSetMode(int mode);             // ThemeMode; applies + notifies windows
int  ThemeModeNow();
void ThemeRefreshDpi(int dpi);
HBRUSH ThemeCtlColorBg(HDC hdc);         // for WM_CTLCOLORSTATIC: sets text/bg, returns bg brush
void ThemeDarkTitlebar(HWND hwnd, bool dark);
void ThemeDarkControls(HWND hwnd, bool dark);   // uxtheme immersive dark for comctl children
// MDL2 glyphs
#define GLYPH_REFRESH      0xE72C
#define GLYPH_CANCEL       0xE711
#define GLYPH_DELETE       0xE74D
#define GLYPH_PAUSE        0xE769
#define GLYPH_PLAY         0xE768
#define GLYPH_RESTART      0xE72C
#define GLYPH_SETTINGS     0xE713
#define GLYPH_INFO         0xE946
#define GLYPH_HELP         0xE897
#define GLYPH_DOWNLOAD     0xE896
#define GLYPH_PRINT        0xE749
#define GLYPH_FOLDER       0xE8B7
#define GLYPH_CHECK        0xE73E
#define GLYPH_CHEVRON      0xE70D
#define GLYPH_SUN          0xE946
#define GLYPH_MOON         0xE708
#define GLYPH_WARNING      0xE7BA
#define GLYPH_SHIELD       0xE7EF
#define GLYPH_LINK         0xE71B
#define GLYPH_COPY         0xE8C8
#define GLYPH_BACK         0xE72B
#define GLYPH_TRAY         0xE786
