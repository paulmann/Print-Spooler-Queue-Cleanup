// Entry point: CLI modes, single instance, first-run wizard, message loop.
#include "app.h"
#include "util.h"
#include "settings.h"
#include "i18n.h"
#include "theme.h"
#include "spooler.h"
#include "main_window.h"
#include "tray.h"
#include "wizard.h"
#include "update.h"
#include "settings_dialog.h"
#include "about.h"

AppState g;
namespace util { bool g_portable_settings = false; }
HINSTANCE g_hInstForControls() { return g.hInst; }

namespace {

void ConsoleWrite(const std::wstring& s) {
    HANDLE h = GetStdHandle(STD_OUTPUT_HANDLE);
    if (h == INVALID_HANDLE_VALUE) return;
    DWORD written = 0;
    std::wstring out = s + L"\n";
    // console expects the active codepage; convert wide -> console CP
    UINT cp = GetConsoleOutputCP();
    if (!cp) cp = CP_ACP;
    int n = WideCharToMultiByte(cp, 0, out.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::string buf((size_t)n, '\0');
    WideCharToMultiByte(cp, 0, out.c_str(), -1, buf.data(), n, nullptr, nullptr);
    WriteFile(h, buf.c_str(), (DWORD)buf.size() - 1, &written, nullptr);
}

bool AttachParentConsole() {
    if (AttachConsole(ATTACH_PARENT_PROCESS)) {
        freopen("CONOUT$", "w", stdout);
        return true;
    }
    return false;
}

void PrintUsage() {
    ConsoleWrite(L"Print Spooler Queue Cleanup " PSQC_VERSION_STR_W);
    ConsoleWrite(L"Usage:");
    ConsoleWrite(L"  SpoolerCleanup.exe                 start the GUI (wizard on first run)");
    ConsoleWrite(L"  SpoolerCleanup.exe /purge          clear all print queues, then exit");
    ConsoleWrite(L"  SpoolerCleanup.exe /list           print current jobs table");
    ConsoleWrite(L"  SpoolerCleanup.exe /status         print spooler service state");
    ConsoleWrite(L"  SpoolerCleanup.exe /tray           start minimized to tray");
    ConsoleWrite(L"  SpoolerCleanup.exe /portable       keep settings next to the exe");
    ConsoleWrite(L"  SpoolerCleanup.exe /install [/allusers|/currentuser] [/dir=PATH]");
    ConsoleWrite(L"                     [/nodesktop] [/nostartmenu] [/noautostart] [/notray] [/silent]");
    ConsoleWrite(L"  SpoolerCleanup.exe /uninstall [/silent]");
    ConsoleWrite(L"  SpoolerCleanup.exe /lang=<tag> /theme=<system|light|dark>");
    ConsoleWrite(L"  SpoolerCleanup.exe /help           this text");
}

int CliPurge() {
    spooler::PurgeResult res;
    spooler::PurgeAll(S().backupBeforePurge, res);
    if (res.ok) {
        ConsoleWrite(util::Fmt(L"OK: removed %d job(s)%s", res.removedBefore,
                   res.backupPath.empty() ? L"" : (L"; backup: " + res.backupPath).c_str()));
        return 0;
    }
    ConsoleWrite(util::Fmt(L"FAIL: %s (code %lu)", res.error.c_str(), res.errCode));
    return (res.errCode == ERROR_ACCESS_DENIED || res.errCode == ERROR_PRIVILEGE_NOT_HELD) ? 2 : 1;
}

int CliList() {
    std::vector<spooler::PrinterJobs> all;
    spooler::EnumerateAll(all);
    int n = 0;
    for (auto& p : all) {
        if (p.jobs.empty()) continue;
        ConsoleWrite(util::Fmt(L"[%s]", p.name.c_str()));
        for (auto& j : p.jobs) {
            ++n;
            ConsoleWrite(util::Fmt(L"  #%lu  %s  %s  %s  %s", j.jobId, j.document.c_str(), j.owner.c_str(),
                       util::FormatBytes(j.size).c_str(), psqc::Str(spooler::JobStatusStrId(j.status))));
        }
    }
    if (!n) ConsoleWrite(L"(empty)");
    return 0;
}

int CliStatus() {
    DWORD st = spooler::ServiceState();
    const wchar_t* s = st == SERVICE_RUNNING ? L"running" :
                       st == SERVICE_STOPPED ? L"stopped" :
                       st ? L"pending" : L"unknown";
    std::vector<spooler::PrinterJobs> all;
    spooler::EnumerateAll(all);
    int n = 0, printers = 0;
    for (auto& p : all) { if (!p.jobs.empty()) ++printers; n += (int)p.jobs.size(); }
    ConsoleWrite(util::Fmt(L"spooler=%s jobs=%d printers_with_jobs=%d version=" PSQC_VERSION_STR_W, s, n, printers));
    return st == SERVICE_RUNNING ? 0 : 1;
}

} // namespace

static bool g_pendingUninstall = false;

int WINAPI wWinMain(HINSTANCE hInst, HINSTANCE, LPWSTR, int showCmd) {
    g.hInst = hInst;
    g.elevated = util::IsElevated();
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);

    int argc = 0;
    LPWSTR* argv = CommandLineToArgvW(GetCommandLineW(), &argc);

    bool startTray = false;
    bool cli = false;
    int cliCode = 0;
    std::wstring langOverride, themeOverride;
    std::vector<std::wstring> args;
    for (int i = 1; i < argc; ++i) args.emplace_back(argv[i]);

    for (auto& a : args) {
        if (util::IEquals(a, L"/portable") || util::FileExists(util::JoinPath(util::ExeDir(), L"portable.txt")))
            util::g_portable_settings = true;
    }
    g.portable = util::g_portable_settings;
    SettingsLoad();
    SetLanguage(S().lang.c_str());
    util::LogInit();
    util::SetLogLevel(S().logLevel);

    for (size_t i = 0; i < args.size(); ++i) {
        const std::wstring& a = args[i];
        if (util::IEquals(a, L"/help") || util::IEquals(a, L"/?") || util::IEquals(a, L"-h")) {
            cli = true; AttachParentConsole(); PrintUsage(); cliCode = 0;
        } else if (util::IEquals(a, L"/purge")) {
            cli = true; AttachParentConsole(); cliCode = CliPurge();
        } else if (util::IEquals(a, L"/list")) {
            cli = true; AttachParentConsole(); cliCode = CliList();
        } else if (util::IEquals(a, L"/status")) {
            cli = true; AttachParentConsole(); cliCode = CliStatus();
        } else if (util::IEquals(a, L"/install")) {
            cli = true; cliCode = CliInstall(argc - (int)i - 1, &argv[i]);
            break;
        } else if (util::IEquals(a, L"/uninstall")) {
            cli = true;
            bool silent = false;
            for (size_t k = i + 1; k < args.size(); ++k) if (util::IEquals(args[k], L"/silent")) silent = true;
            // needs windows for confirm box unless silent
            if (silent) { CliUninstall(true); cliCode = 0; }
            else { cli = false; g.quitRequested = true; /* GUI path handles below */ g_pendingUninstall = true; }
        } else if (util::IEquals(a, L"/tray") || util::IEquals(a, L"/minimized")) {
            startTray = true;
        } else if (util::StartsWithI(a, L"/lang=")) {
            langOverride = a.substr(6);
        } else if (util::StartsWithI(a, L"/theme=")) {
            themeOverride = a.substr(7);
        } else if (util::IEquals(a, L"/replace")) {
            cli = true;
            if (i + 3 < args.size())
                cliCode = update::RunSelfReplace(args[i + 1], args[i + 2], args[i + 3]);
            else cliCode = 1;
            break;
        }
    }
    if (cli) {
        LocalFree(argv);
        util::LogClose();
        CoUninitialize();
        return cliCode;
    }
    if (!langOverride.empty()) { S().lang = langOverride; SetLanguage(S().lang.c_str()); }
    if (!themeOverride.empty()) {
        int m = util::IEquals(themeOverride, L"light") ? THEME_LIGHT :
                util::IEquals(themeOverride, L"dark") ? THEME_DARK : THEME_SYSTEM;
        S().themeMode = m;
    }

    // single instance
    HANDLE mutex = CreateMutexW(nullptr, FALSE, PSQC_MUTEX_NAME);
    if (GetLastError() == ERROR_ALREADY_EXISTS && mutex) {
        HWND other = FindWindowW(PSQC_WINDOW_CLASS, nullptr);
        if (other) { PostMessageW(other, WM_PSQC_SINGLETON, 0, 0); SetForegroundWindow(other); }
        CloseHandle(mutex);
        CoUninitialize();
        return 0;
    }

    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_LISTVIEW_CLASSES | ICC_LINK_CLASS | ICC_STANDARD_CLASSES };
    InitCommonControlsEx(&icc);

    ThemeInit(util::DpiOf(nullptr));

    LOGA(L"start: version=" PSQC_VERSION_STR_W L" elevated=%d portable=%d lang=%s",
         g.elevated ? 1 : 0, g.portable ? 1 : 0, Langs()[CurrentLang()].tag);

    // first-run wizard
    if (!S().firstRunDone && !g.portable) {
        WizardResult wr = WizardRun();
        if (wr == WR_Cancelled) {
            CloseHandle(mutex);
            util::LogClose();
            CoUninitialize();
            return 0;
        }
    } else if (!S().firstRunDone) {
        S().firstRunDone = true;
        SettingsSave();
    }

    int show = startTray ? SW_HIDE : showCmd;
    if (S().startMinimized) show = SW_HIDE;
    MainWindowCreate(hInst, show);
    if (S().trayIcon) TrayCreate();
    if (g_pendingUninstall) { UninstallStart(false); }
    if (startTray && S().trayIcon) {
        // stay hidden; user opens via tray dblclick
    }

    if (getenv("PSQC_SHOTDIR")) {          // QA self-capture drive: open every dialog once
        SettingsShow();
        AboutShow();
        PostQuitMessage(0);
    } else if (S().autoUpdate && util::UnixSeconds() - S().lastUpdateCheck > 86400)
        UpdateCheckAsync(false);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageW(&msg);
    }

    if (g.timerRefresh) KillTimer(g.hwndMain, 1);
    LOGA(L"exit");
    util::LogClose();
    CloseHandle(mutex);
    CoUninitialize();
    return (int)msg.wParam;
}
