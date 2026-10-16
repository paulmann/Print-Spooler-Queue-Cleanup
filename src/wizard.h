#pragma once
// Built-in installer: first-run wizard, install/uninstall, shortcuts, autostart,
// Programs & Features registration. No external installer framework.
#include "app.h"

enum WizardResult : int { WR_Cancelled = 0, WR_Installed = 1, WR_RunOnce = 2 };

struct InstallOptions {
    int scope = 0;                 // 0 current user, 1 all users
    std::wstring dir;
    bool desktop = true;
    bool startMenu = true;
    bool autostart = true;
    bool trayAtLogon = true;
    bool silent = false;
};

WizardResult WizardRun();                       // modal first-run wizard
bool InstallExecute(const InstallOptions& o, std::wstring& err);
void UninstallStart(bool silent);               // confirm (unless silent), remove, schedule self-delete, exit
std::wstring DefaultInstallDir(int scope);
// CLI helpers
int CliInstall(int argc, wchar_t** argv);       // parses /allusers /dir=... /silent etc.
int CliUninstall(bool silent);
