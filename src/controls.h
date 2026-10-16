#pragma once
// Custom-drawn controls: Fluent-style buttons, switches, status chips.
// Both themes, DPI-aware, keyboard accessible (focus rects), tooltips.
#include "app.h"

enum CtrlKind : int { CK_Button = 0, CK_Toggle = 1, CK_Chip = 2 };
enum BtnKind : int { BK_Standard = 0, BK_Accent = 1, BK_Danger = 2, BK_Subtle = 3 };

HWND CreatePsqcButton(HWND parent, int ctrlId, int strId, wchar_t glyph, int btnKind);
HWND CreatePsqcToggle(HWND parent, int ctrlId, int strId, bool checked);   // switch + optional label
HWND CreatePsqcChip(HWND parent, int ctrlId);                              // status pill, set text later
void   CtrlSetChecked(HWND h, bool on);
bool   CtrlChecked(HWND h);
void   CtrlSetStrId(HWND h, int strId);
void   CtrlSetText(HWND h, const std::wstring& text);       // overrides strId (chips, dynamic labels)
std::wstring CtrlText(HWND h);
void   CtrlSetGlyph(HWND h, wchar_t glyph);
void   CtrlSetKind(HWND h, int btnKind);
void   CtrlSetDotColor(HWND h, COLORREF c);                 // chip dot
void   CtrlRelabelAll();                                    // after language change
void   CtrlAddTip(HWND h, int strId);
void   CtrlSetEnabledVisual(HWND h, bool enabled);
bool   CtrlHandleDrawItem(DRAWITEMSTRUCT* dis);             // parent WM_DRAWITEM
void   CtrlResizeToContent(HWND h, HDC hdcOpt);             // width = glyph+text+padding
