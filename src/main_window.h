#pragma once
// Main application window: header, toolbar, job list, settings row, status bar.
#include "app.h"
#include "update.h"

enum MwIds : int {
    IDC_BTN_REFRESH = 101, IDC_BTN_DELSEL, IDC_BTN_CLEARALL, IDC_BTN_PAUSE, IDC_BTN_RESTART,
    IDC_BTN_SETTINGS, IDC_BTN_THEME, IDC_BTN_ABOUT,
    IDC_LIST, IDC_EMPTY_TITLE, IDC_EMPTY_HINT,
    IDC_CMB_LANG, IDC_CMB_THEME, IDC_TGL_TRAY, IDC_TGL_AUTOSTART,
    IDC_LBL_LANG, IDC_LBL_THEME, IDC_LBL_TRAY, IDC_LBL_AUTOSTART,
};

void MainWindowCreate(HINSTANCE hInst, int showCmd);
void MainWindowRefreshNow();
void MainWindowReskin();            // theme and/or language changed: full re-skin
void MainWindowShowAndActivate();
void MainWindowBeginPurge(bool confirm);
bool MainWindowBusy();
void MainWindowSetJobCount(int n);  // updates chip + cached count for tray
HWND MainWindowHandle();
void UpdateCheckAsync(bool manual);             // GitHub releases check in background
void UpdateApplyAsync(const update::ReleaseInfo* info);  // download + self-replace
struct UpdateCheckUi { bool newer = false; bool manual = false; std::wstring err; update::ReleaseInfo info; };
struct UpdateDlUi { bool ok = false; std::wstring path; std::wstring err; };

