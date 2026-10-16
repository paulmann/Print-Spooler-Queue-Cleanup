#pragma once
// Tray icon: context menu (spec + extras), live job counter, double-click, balloons.
#include "app.h"

void TrayCreate();
void TrayDestroy();
bool TrayPresent();
void TrayUpdateCounts(int jobs);                 // menu text + tooltip
void TrayNotify(int kind, const std::wstring& text);
void TrayReskin();                               // theme/language changed
