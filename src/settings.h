#pragma once
// Persistent user settings (JSON in %APPDATA%\SpoolerCleanup, or next to exe in portable mode).
#include "app.h"

enum ThemeMode : int { THEME_SYSTEM = 0, THEME_LIGHT = 1, THEME_DARK = 2 };

struct Settings {
    int         themeMode         = THEME_SYSTEM;
    std::wstring lang;                       // "" = follow system; else BCP-ish tag from i18n table
    bool        trayIcon          = true;
    bool        minimizeToTray    = true;
    bool        startMinimized    = false;
    bool        notifications     = true;
    bool        confirmClearAll   = true;
    bool        autoUpdate        = true;
    bool        autoElevate       = true;
    bool        backupBeforePurge = true;
    int         logLevel          = 1;       // util::LogLevel
    int         winX = -1, winY = -1, winW = -1, winH = -1;
    bool        winMaximized      = false;
    uint64_t    lastUpdateCheck   = 0;
    std::wstring skipVersion;
    bool        firstRunDone      = false;
    int         installScope      = 0;       // 0 = current user, 1 = all users
    std::wstring installPath;
    int         refreshSec        = 3;       // live list refresh period, 0 = off
};

Settings& S();
void SettingsLoad();
void SettingsSave();

// "Start with Windows" (registry Run). Scope follows install scope; all-users needs admin.
bool AutostartEnabled();
bool AutostartSet(bool on, std::wstring& err);
std::wstring AutostartRunValue();
