// Persistent settings implementation.
#include "settings.h"
#include "util.h"

namespace util { extern void SetLogLevel(int lvl); }

static Settings g_settings;
Settings& S() { return g_settings; }

void SettingsLoad() {
    std::string raw;
    if (!util::ReadFileAll(util::SettingsFilePath(), raw)) return;
    util::JsonValue root;
    if (!util::JsonParse(raw, root)) return;
    auto GetB = [&](const wchar_t* k, bool& out) { if (auto* v = root.Find(k)) if (v->type == util::JsonValue::Bool) out = v->b; };
    auto GetI = [&](const wchar_t* k, int& out) { if (auto* v = root.Find(k)) if (v->type == util::JsonValue::Num) out = (int)v->num; };
    auto GetS = [&](const wchar_t* k, std::wstring& out) { if (auto* v = root.Find(k)) if (v->type == util::JsonValue::Str) out = v->str; };
    auto GetU = [&](const wchar_t* k, uint64_t& out) { if (auto* v = root.Find(k)) if (v->type == util::JsonValue::Num) out = (uint64_t)v->num; };
    GetI(L"themeMode", g_settings.themeMode);
    GetS(L"lang", g_settings.lang);
    GetS(L"fontFace", g_settings.fontFace);
    GetI(L"fontSize", g_settings.fontSize);
    if (g_settings.fontSize < 8 || g_settings.fontSize > 16) g_settings.fontSize = 10;
    GetB(L"trayIcon", g_settings.trayIcon);
    GetB(L"minimizeToTray", g_settings.minimizeToTray);
    GetB(L"startMinimized", g_settings.startMinimized);
    GetB(L"notifications", g_settings.notifications);
    GetB(L"confirmClearAll", g_settings.confirmClearAll);
    GetB(L"autoUpdate", g_settings.autoUpdate);
    GetB(L"autoElevate", g_settings.autoElevate);
    GetB(L"backupBeforePurge", g_settings.backupBeforePurge);
    GetI(L"logLevel", g_settings.logLevel);
    GetI(L"winX", g_settings.winX);
    GetI(L"winY", g_settings.winY);
    GetI(L"winW", g_settings.winW);
    GetI(L"winH", g_settings.winH);
    GetB(L"winMaximized", g_settings.winMaximized);
    GetU(L"lastUpdateCheck", g_settings.lastUpdateCheck);
    GetS(L"skipVersion", g_settings.skipVersion);
    GetB(L"firstRunDone", g_settings.firstRunDone);
    GetI(L"installScope", g_settings.installScope);
    GetS(L"installPath", g_settings.installPath);
    GetI(L"refreshSec", g_settings.refreshSec);
    if (g_settings.themeMode < 0 || g_settings.themeMode > 2) g_settings.themeMode = THEME_SYSTEM;
    if (g_settings.refreshSec < 0 || g_settings.refreshSec > 60) g_settings.refreshSec = 3;
    util::SetLogLevel(g_settings.logLevel);
}

void SettingsSave() {
    auto B = [](const char* k, bool v) { return std::string("\"") + k + "\": " + (v ? "true" : "false"); };
    auto I = [](const char* k, int v) { return std::string("\"") + k + "\": " + std::to_string(v); };
    auto U = [](const char* k, uint64_t v) { return std::string("\"") + k + "\": " + std::to_string(v); };
    auto Str = [](const char* k, const std::wstring& v) {
        return std::string("\"") + k + "\": \"" + util::JsonEscape(v) + "\"";
    };
    std::string j = "{\n";
    j += "  \"schema\": 1,\n  ";
    j += Str("product", PSQC_VERSION_STR_W) + ",\n  ";
    j += I("themeMode", g_settings.themeMode) + ",\n  ";
    j += Str("lang", g_settings.lang) + ",\n  ";
    j += Str("fontFace", g_settings.fontFace) + ",\n  ";
    j += I("fontSize", g_settings.fontSize) + ",\n  ";
    j += B("trayIcon", g_settings.trayIcon) + ",\n  ";
    j += B("minimizeToTray", g_settings.minimizeToTray) + ",\n  ";
    j += B("startMinimized", g_settings.startMinimized) + ",\n  ";
    j += B("notifications", g_settings.notifications) + ",\n  ";
    j += B("confirmClearAll", g_settings.confirmClearAll) + ",\n  ";
    j += B("autoUpdate", g_settings.autoUpdate) + ",\n  ";
    j += B("autoElevate", g_settings.autoElevate) + ",\n  ";
    j += B("backupBeforePurge", g_settings.backupBeforePurge) + ",\n  ";
    j += I("logLevel", g_settings.logLevel) + ",\n  ";
    j += I("winX", g_settings.winX) + ",\n  ";
    j += I("winY", g_settings.winY) + ",\n  ";
    j += I("winW", g_settings.winW) + ",\n  ";
    j += I("winH", g_settings.winH) + ",\n  ";
    j += B("winMaximized", g_settings.winMaximized) + ",\n  ";
    j += U("lastUpdateCheck", g_settings.lastUpdateCheck) + ",\n  ";
    j += Str("skipVersion", g_settings.skipVersion) + ",\n  ";
    j += B("firstRunDone", g_settings.firstRunDone) + ",\n  ";
    j += I("installScope", g_settings.installScope) + ",\n  ";
    j += Str("installPath", g_settings.installPath) + ",\n  ";
    j += I("refreshSec", g_settings.refreshSec) + "\n}\n";
    util::WriteFileAll(util::SettingsFilePath(), j);
}

static const wchar_t* RUN_KEY = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
static const wchar_t* RUN_VALUE = L"SpoolerCleanup";

std::wstring AutostartRunValue() {
    return L"\"" + util::ExePath() + L"\" /tray";
}

bool AutostartEnabled() {
    std::wstring v;
    if (util::RegGetString(HKEY_CURRENT_USER, RUN_KEY, RUN_VALUE, v)) return true;
    if (g_settings.installScope == 1 && util::RegGetString(HKEY_LOCAL_MACHINE, RUN_KEY, RUN_VALUE, v)) return true;
    return false;
}

bool AutostartSet(bool on, std::wstring& err) {
    bool ok = true;
    if (on) {
        if (g_settings.installScope == 1 && util::IsElevated()) {
            ok = util::RegSetString(HKEY_LOCAL_MACHINE, RUN_KEY, RUN_VALUE, AutostartRunValue().c_str());
        } else {
            ok = util::RegSetString(HKEY_CURRENT_USER, RUN_KEY, RUN_VALUE, AutostartRunValue().c_str());
        }
        // clear the other scope to avoid double entries
        util::RegDeleteValue(g_settings.installScope == 1 && util::IsElevated() ? HKEY_CURRENT_USER : HKEY_LOCAL_MACHINE, RUN_KEY, RUN_VALUE);
    } else {
        bool a = util::RegDeleteValue(HKEY_CURRENT_USER, RUN_KEY, RUN_VALUE);
        bool b = util::RegDeleteValue(HKEY_LOCAL_MACHINE, RUN_KEY, RUN_VALUE);
        ok = a || b;
    }
    if (!ok) err = util::LastErrorText(GetLastError());
    return ok;
}
