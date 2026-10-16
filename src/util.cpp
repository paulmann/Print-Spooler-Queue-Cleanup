// Utility layer implementation.
#include "util.h"
#include <shlwapi.h>
#include <sddl.h>

bool RtlActiveViaI18n();   // defined in i18n.cpp (global namespace)

namespace util {

// ---- strings ---------------------------------------------------------------
std::wstring Utf8ToWide(std::string_view utf8) {
    if (utf8.empty()) return {};
    int n = MultiByteToWideChar(CP_UTF8, 0, utf8.data(), (int)utf8.size(), nullptr, 0);
    std::wstring out((size_t)n, L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8.data(), (int)utf8.size(), out.data(), n);
    return out;
}

std::string WideToUtf8(std::wstring_view wide) {
    if (wide.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, wide.data(), (int)wide.size(), nullptr, 0, nullptr, nullptr);
    std::string out((size_t)n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, wide.data(), (int)wide.size(), out.data(), n, nullptr, nullptr);
    return out;
}

std::wstring Fmt(const wchar_t* fmt, ...) {
    va_list ap;
    va_start(ap, fmt);
    int n = _vscwprintf(fmt, ap);
    va_end(ap);
    if (n < 0) return {};
    std::wstring out((size_t)n, L'\0');
    va_start(ap, fmt);
    vswprintf_s(out.data(), out.size() + 1, fmt, ap);
    va_end(ap);
    return out;
}

std::wstring Trim(std::wstring_view s) {
    size_t b = s.find_first_not_of(L" \t\r\n");
    if (b == std::wstring_view::npos) return {};
    size_t e = s.find_last_not_of(L" \t\r\n");
    return std::wstring(s.substr(b, e - b + 1));
}

std::wstring Lower(std::wstring_view s) {
    std::wstring out(s);
    for (auto& c : out) c = (wchar_t)towlower((wchar_t)c);
    return out;
}

bool IEquals(std::wstring_view a, std::wstring_view b) {
    if (a.size() != b.size()) return false;
    return CompareStringOrdinal(a.data(), (int)a.size(), b.data(), (int)b.size(), TRUE) == CSTR_EQUAL;
}

bool StartsWithI(std::wstring_view s, std::wstring_view p) {
    return s.size() >= p.size() && IEquals(s.substr(0, p.size()), p);
}

bool EndsWithI(std::wstring_view s, std::wstring_view suf) {
    return s.size() >= suf.size() && IEquals(s.substr(s.size() - suf.size()), suf);
}

std::vector<std::wstring> Split(std::wstring_view s, wchar_t sep) {
    std::vector<std::wstring> out;
    size_t start = 0;
    while (true) {
        size_t p = s.find(sep, start);
        if (p == std::wstring_view::npos) { out.emplace_back(s.substr(start)); break; }
        out.emplace_back(s.substr(start, p - start));
        start = p + 1;
    }
    return out;
}

std::wstring FormatBytes(uint64_t bytes) {
    const wchar_t* units[] = { L"B", L"KB", L"MB", L"GB", L"TB" };
    double v = (double)bytes;
    int u = 0;
    while (v >= 1024.0 && u < 4) { v /= 1024.0; ++u; }
    if (u == 0) return Fmt(L"%llu B", (unsigned long long)bytes);
    return Fmt(L"%.2f %s", v, units[u]);
}

std::wstring FormatInt(uint64_t v) { return Fmt(L"%llu", (unsigned long long)v); }

// ---- paths -----------------------------------------------------------------
std::wstring ExePath() {
    wchar_t buf[MAX_PATH + 1]{};
    DWORD n = GetModuleFileNameW(nullptr, buf, MAX_PATH);
    return std::wstring(buf, n);
}

std::wstring ExeDir() {
    auto p = ExePath();
    size_t i = p.find_last_of(L"\\/");
    return i == std::wstring::npos ? std::wstring() : p.substr(0, i);
}

static std::wstring KnownFolder(REFKNOWNFOLDERID id, bool create) {
    PWSTR raw = nullptr;
    HRESULT hr = SHGetKnownFolderPath(id, create ? 0 : KF_FLAG_DONT_VERIFY, nullptr, &raw);
    std::wstring out;
    if (SUCCEEDED(hr) && raw) out = raw;
    if (raw) CoTaskMemFree(raw);
    return out;
}

std::wstring AppDataDir(bool create) {
    auto d = JoinPath(KnownFolder(FOLDERID_RoamingAppData, true), PSQC_SETTINGS_DIR);
    if (create) EnsureDir(d);
    return d;
}

std::wstring LocalDataDir(bool create) {
    auto d = JoinPath(KnownFolder(FOLDERID_LocalAppData, true), PSQC_SETTINGS_DIR);
    if (create) EnsureDir(d);
    return d;
}

std::wstring LogsDir(bool create) {
    auto d = JoinPath(LocalDataDir(create), L"logs");
    if (create) EnsureDir(d);
    return d;
}

std::wstring BackupRootDir(bool create) {
    auto d = JoinPath(LocalDataDir(create), L"backup");
    if (create) EnsureDir(d);
    return d;
}

extern bool g_portable_settings;   // set by main.cpp from /portable or portable.txt
std::wstring SettingsFilePath() {
    if (g_portable_settings) return JoinPath(ExeDir(), L"settings.json");
    return JoinPath(AppDataDir(true), L"settings.json");
}

std::wstring JoinPath(std::wstring_view dir, std::wstring_view file) {
    if (dir.empty()) return std::wstring(file);
    if (file.empty()) return std::wstring(dir);
    wchar_t last = dir.back();
    if (last == L'\\' || last == L'/') return std::wstring(dir) + std::wstring(file);
    return std::wstring(dir) + L"\\" + std::wstring(file);
}

std::wstring SpoolDir() {
    wchar_t buf[MAX_PATH + 1]{};
    GetSystemDirectoryW(buf, MAX_PATH);   // C:\Windows\System32
    return JoinPath(std::wstring(buf), L"spool\\PRINTERS");
}

std::wstring TempDir() {
    wchar_t buf[MAX_PATH + 1]{};
    DWORD n = GetTempPathW(MAX_PATH, buf);
    return std::wstring(buf, n);
}

// ---- files -----------------------------------------------------------------
bool FileExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && !(a & FILE_ATTRIBUTE_DIRECTORY);
}

bool DirExists(const std::wstring& p) {
    DWORD a = GetFileAttributesW(p.c_str());
    return a != INVALID_FILE_ATTRIBUTES && (a & FILE_ATTRIBUTE_DIRECTORY);
}

bool EnsureDir(const std::wstring& p) {
    if (p.empty() || DirExists(p)) return DirExists(p);
    size_t i = p.find_last_of(L"\\/");
    if (i != std::wstring::npos && i + 1 < p.size()) EnsureDir(p.substr(0, i));
    return CreateDirectoryW(p.c_str(), nullptr) || GetLastError() == ERROR_ALREADY_EXISTS;
}

bool ReadFileAll(const std::wstring& p, std::string& out) {
    HANDLE h = CreateFileW(p.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    LARGE_INTEGER sz{};
    GetFileSizeEx(h, &sz);
    out.resize((size_t)sz.QuadPart);
    DWORD read = 0;
    bool ok = sz.QuadPart == 0 || ReadFile(h, out.data(), (DWORD)out.size(), &read, nullptr);
    CloseHandle(h);
    return ok;
}

bool WriteFileAll(const std::wstring& p, std::string_view data) {
    HANDLE h = CreateFileW(p.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    DWORD written = 0;
    bool ok = data.empty() || WriteFile(h, data.data(), (DWORD)data.size(), &written, nullptr);
    CloseHandle(h);
    return ok && written == data.size();
}

bool RemoveFile(const std::wstring& p) {
    SetFileAttributesW(p.c_str(), FILE_ATTRIBUTE_NORMAL);
    return DeleteFileW(p.c_str()) != 0;
}

bool RemoveDirTree(const std::wstring& p) {
    std::wstring pattern = JoinPath(p, L"*");
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(pattern.c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return false;
    do {
        if (wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        auto full = JoinPath(p, fd.cFileName);
        if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) RemoveDirTree(full);
        else RemoveFile(full);
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return RemoveDirectoryW(p.c_str()) != 0;
}

uint64_t FileSize(const std::wstring& p) {
    WIN32_FILE_ATTRIBUTE_DATA fa{};
    if (!GetFileAttributesExW(p.c_str(), GetFileExInfoStandard, &fa)) return 0;
    return ((uint64_t)fa.nFileSizeHigh << 32) | fa.nFileSizeLow;
}

// ---- registry --------------------------------------------------------------
static LSTATUS OpenKey(HKEY root, const wchar_t* subkey, REGSAM sam, HKEY& out, bool create) {
    if (create) return RegCreateKeyExW(root, subkey, 0, nullptr, 0, sam, nullptr, &out, nullptr);
    return RegOpenKeyExW(root, subkey, 0, sam, &out);
}

bool RegGetDword(HKEY root, const wchar_t* subkey, const wchar_t* name, DWORD& out) {
    HKEY k{};
    if (OpenKey(root, subkey, KEY_READ, k, false) != ERROR_SUCCESS) return false;
    DWORD type = 0, size = sizeof(out);
    LSTATUS st = RegQueryValueExW(k, name, nullptr, &type, (BYTE*)&out, &size);
    RegCloseKey(k);
    return st == ERROR_SUCCESS && (type == REG_DWORD || type == REG_SZ);
}

bool RegSetDword(HKEY root, const wchar_t* subkey, const wchar_t* name, DWORD value) {
    HKEY k{};
    if (OpenKey(root, subkey, KEY_WRITE, k, true) != ERROR_SUCCESS) return false;
    LSTATUS st = RegSetValueExW(k, name, 0, REG_DWORD, (const BYTE*)&value, sizeof(value));
    RegCloseKey(k);
    return st == ERROR_SUCCESS;
}

bool RegGetString(HKEY root, const wchar_t* subkey, const wchar_t* name, std::wstring& out) {
    HKEY k{};
    if (OpenKey(root, subkey, KEY_READ, k, false) != ERROR_SUCCESS) return false;
    DWORD type = 0, size = 0;
    if (RegQueryValueExW(k, name, nullptr, &type, nullptr, &size) != ERROR_SUCCESS || (type != REG_SZ && type != REG_EXPAND_SZ)) {
        RegCloseKey(k);
        return false;
    }
    out.resize(size / sizeof(wchar_t) + 1, L'\0');
    LSTATUS st = RegQueryValueExW(k, name, nullptr, &type, (BYTE*)out.data(), &size);
    RegCloseKey(k);
    if (st != ERROR_SUCCESS) return false;
    while (!out.empty() && out.back() == L'\0') out.pop_back();
    return true;
}

bool RegSetString(HKEY root, const wchar_t* subkey, const wchar_t* name, const wchar_t* value) {
    HKEY k{};
    if (OpenKey(root, subkey, KEY_WRITE, k, true) != ERROR_SUCCESS) return false;
    LSTATUS st = RegSetValueExW(k, name, 0, REG_SZ, (const BYTE*)value, (DWORD)((wcslen(value) + 1) * sizeof(wchar_t)));
    RegCloseKey(k);
    return st == ERROR_SUCCESS;
}

bool RegDeleteValue(HKEY root, const wchar_t* subkey, const wchar_t* name) {
    HKEY k{};
    if (OpenKey(root, subkey, KEY_WRITE, k, false) != ERROR_SUCCESS) return false;
    LSTATUS st = RegDeleteValueW(k, name);
    RegCloseKey(k);
    return st == ERROR_SUCCESS;
}

bool RegDeleteTree(HKEY root, const wchar_t* subkey) {
    return SHDeleteKeyW(root, subkey) == ERROR_SUCCESS;
}

// ---- system ----------------------------------------------------------------
bool IsElevated() {
    HANDLE tok{};
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &tok)) return false;
    TOKEN_ELEVATION el{};
    DWORD n = 0;
    bool ok = GetTokenInformation(tok, TokenElevation, &el, sizeof(el), &n) && el.TokenIsElevated;
    CloseHandle(tok);
    return ok;
}

DWORD WindowsBuild() {
    DWORD v = 0;
    if (RegGetDword(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", L"CurrentBuildNumber", v))
        return v;
    return 0;
}

std::wstring OsFriendlyName() {
    std::wstring edition;
    RegGetString(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", L"EditionID", edition);
    DWORD major = 0, minor = 0;
    RegGetDword(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", L"CurrentMajorVersionNumber", major);
    RegGetDword(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion", L"CurrentMinorVersionNumber", minor);
    return Fmt(L"Windows %lu.%lu (%s), build %lu", major, minor,
               edition.empty() ? L"edition?" : edition.c_str(), WindowsBuild());
}

void OpenUrl(const char* url) {
    ShellExecuteA(nullptr, "open", url, nullptr, nullptr, SW_SHOWNORMAL);
}

bool ShellRun(const wchar_t* verb, const wchar_t* file, const wchar_t* params, bool wait, DWORD* exitCode) {
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC | (wait ? SEE_MASK_NOCLOSEPROCESS : 0);
    sei.lpVerb = verb;
    sei.lpFile = file;
    sei.lpParameters = params;
    sei.nShow = SW_SHOWNORMAL;
    if (!ShellExecuteExW(&sei)) return false;
    if (wait && sei.hProcess) {
        WaitForSingleObject(sei.hProcess, INFINITE);
        if (exitCode) GetExitCodeProcess(sei.hProcess, exitCode);
        CloseHandle(sei.hProcess);
    }
    return true;
}

bool RunElevatedAndWait(const wchar_t* file, const wchar_t* params, DWORD* exitCode) {
    return ShellRun(L"runas", file, params, true, exitCode);
}

std::wstring DefaultPrinterName() {
    DWORD n = 0;
    GetDefaultPrinterW(nullptr, &n);
    if (!n) return {};
    std::wstring out(n, L'\0');
    if (!GetDefaultPrinterW(out.data(), &n)) return {};
    while (!out.empty() && out.back() == L'\0') out.pop_back();
    return out;
}

void OpenPrinterQueueUI(const std::wstring& printerName) {
    if (!printerName.empty()) {
        std::wstring params = Fmt(L"printui.dll,PrintUIEntry /o /n \"%s\"", printerName.c_str());
        if (ShellRun(L"open", L"rundll32.exe", params.c_str(), false, nullptr)) return;
    }
    OpenPrintersFolder();
}

void OpenPrintersFolder() {
    // Classic "Devices and Printers"-era printers folder; on Win11 falls back gracefully.
    if (!ShellRun(L"open", L"explorer.exe", L"shell:::{A8A91A66-3A7D-4424-8D24-04E180695C7A}", false, nullptr))
        OpenPrintersSettingsPage();
}

void OpenPrintersSettingsPage() {
    ShellRun(L"open", L"ms-settings:printers", nullptr, false, nullptr);
}

void OpenFolder(const std::wstring& path) {
    ShellRun(L"open", L"explorer.exe", path.c_str(), false, nullptr);
}

std::wstring LastErrorText(DWORD err) {
    wchar_t* buf = nullptr;
    DWORD n = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM |
                             FORMAT_MESSAGE_IGNORE_INSERTS, nullptr, err,
                             MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPWSTR)&buf, 0, nullptr);
    std::wstring out;
    if (n && buf) out.assign(buf, n);
    if (buf) LocalFree(buf);
    out = Trim(out);
    if (out.empty()) out = Fmt(L"error %lu", err);
    return out;
}

std::wstring CurrentUserSid() {
    HANDLE tok{};
    if (!OpenProcessToken(GetCurrentProcess(), TOKEN_QUERY, &tok)) return {};
    DWORD need = 0;
    GetTokenInformation(tok, TokenUser, nullptr, 0, &need);
    std::vector<BYTE> buf(need);
    if (!GetTokenInformation(tok, TokenUser, buf.data(), need, &need)) { CloseHandle(tok); return {}; }
    CloseHandle(tok);
    auto* tu = (TOKEN_USER*)buf.data();
    PWSTR sid = nullptr;
    if (!ConvertSidToStringSidW(tu->User.Sid, &sid)) return {};
    std::wstring out = sid;
    LocalFree(sid);
    return out;
}

// ---- time ------------------------------------------------------------------
std::wstring FormatSystemTimeUser(const SYSTEMTIME& st) {
    wchar_t d[64]{}, t[64]{};
    GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, DATE_SHORTDATE, &st, nullptr, d, 64, nullptr);
    GetTimeFormatEx(LOCALE_NAME_USER_DEFAULT, TIME_NOSECONDS, &st, nullptr, t, 64);
    return Fmt(L"%s %s", d, t);
}

std::wstring NowStamp() {
    SYSTEMTIME st{};
    GetLocalTime(&st);
    return Fmt(L"%04u%02u%02u-%02u%02u%02u", st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
}

uint64_t UnixSeconds() {
    FILETIME ft{};
    GetSystemTimeAsFileTime(&ft);
    uint64_t t = ((uint64_t)ft.dwHighDateTime << 32) | ft.dwLowDateTime;
    return (t - 116444736000000000ULL) / 10000000ULL;
}

// ---- dpi -------------------------------------------------------------------
int DpiOf(HWND hwnd) {
    UINT d = hwnd ? GetDpiForWindow(hwnd) : 0;
    if (!d) {
        HDC dc = GetDC(nullptr);
        d = (UINT)GetDeviceCaps(dc, LOGPIXELSY);
        ReleaseDC(nullptr, dc);
    }
    return d ? (int)d : 96;
}

int Scale(HWND hwnd, int dip) { return MulDiv(dip, DpiOf(hwnd), 96); }

// ---- misc ------------------------------------------------------------------
static bool RtlActive();
void ApplyRtlLayout(HWND root) {
    if (!RtlActive()) return;
    auto Set = [](HWND h) {
        LONG_PTR ex = GetWindowLongPtrW(h, GWL_EXSTYLE);
        if (!(ex & WS_EX_LAYOUTRTL)) {
            SetWindowLongPtrW(h, GWL_EXSTYLE, ex | WS_EX_LAYOUTRTL);
        }
    };
    Set(root);
    EnumChildWindows(root, [](HWND c, LPARAM) -> BOOL { 
        LONG_PTR ex = GetWindowLongPtrW(c, GWL_EXSTYLE);
        SetWindowLongPtrW(c, GWL_EXSTYLE, ex | WS_EX_LAYOUTRTL);
        return TRUE;
    }, 0);
    SetWindowPos(root, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
}

static bool RtlActive() { return RtlActiveViaI18n(); }

bool CopyToClipboard(const std::wstring& text) {
    if (!OpenClipboard(nullptr)) return false;
    EmptyClipboard();
    HGLOBAL h = GlobalAlloc(GMEM_MOVEABLE, (text.size() + 1) * sizeof(wchar_t));
    if (!h) { CloseClipboard(); return false; }
    wchar_t* p = (wchar_t*)GlobalLock(h);
    memcpy(p, text.data(), text.size() * sizeof(wchar_t));
    p[text.size()] = 0;
    GlobalUnlock(h);
    SetClipboardData(CF_UNICODETEXT, h);
    CloseClipboard();
    return true;
}

HICON LoadAppIcon(int cx, int cy) {
    return (HICON)LoadImageW(GetModuleHandleW(nullptr), MAKEINTRESOURCEW(1), IMAGE_ICON, cx, cy, LR_SHARED);
}

std::wstring LoadStringRes(UINT id) {
    const wchar_t* p = nullptr;
    int n = LoadStringW(GetModuleHandleW(nullptr), id, (LPWSTR)&p, 0);
    return n > 0 ? std::wstring(p, n) : std::wstring();
}

RECT WorkAreaRect() {
    RECT r{ 0, 0, 1280, 800 };
    SystemParametersInfoW(SPI_GETWORKAREA, 0, &r, 0);
    return r;
}

void CenterWindow(HWND hwnd, HWND owner) {
    RECT rc{};
    GetWindowRect(hwnd, &rc);
    RECT base = WorkAreaRect();
    if (owner && IsWindowVisible(owner)) GetWindowRect(owner, &base);
    int x = base.left + ((base.right - base.left) - (rc.right - rc.left)) / 2;
    int y = base.top + ((base.top == 0 ? base.bottom - base.top : base.bottom - base.top) - (rc.bottom - rc.top)) / 2;
    y = base.top + ((base.bottom - base.top) - (rc.bottom - rc.top)) / 2;
    SetWindowPos(hwnd, nullptr, x, y, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
}

// ---- logging ---------------------------------------------------------------
static std::mutex g_logMutex;
static std::wstring g_logPath;
static int g_logLevel = LOG_ERR;
void SetLogLevel(int lvl);
int  GetLogLevel() { return g_logLevel; }
void SetLogLevel(int lvl) { g_logLevel = lvl; }

void LogInit() {
    g_logPath = JoinPath(LogsDir(true), Fmt(L"app-%s.log", NowStamp().substr(0, 8).c_str()));
    // prune old logs (>14 files)
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(JoinPath(LogsDir(false), L"app-*.log").c_str(), &fd);
    std::vector<std::wstring> files;
    if (h != INVALID_HANDLE_VALUE) {
        do files.emplace_back(fd.cFileName); while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    std::sort(files.begin(), files.end());
    while (files.size() > 14) { RemoveFile(JoinPath(LogsDir(false), files.front())); files.erase(files.begin()); }
}

void LogClose() { g_logPath.clear(); }

void Log(int level, const wchar_t* fmt, ...) {
    if (g_logPath.empty() || level > g_logLevel || g_logLevel == LOG_OFF) return;
    va_list ap;
    va_start(ap, fmt);
    int n = _vscwprintf(fmt, ap);
    va_end(ap);
    if (n < 0) return;
    std::wstring msg((size_t)n, L'\0');
    va_start(ap, fmt);
    vswprintf_s(msg.data(), msg.size() + 1, fmt, ap);
    va_end(ap);
    SYSTEMTIME st{};
    GetLocalTime(&st);
    std::wstring line = Fmt(L"[%04u-%02u-%02u %02u:%02u:%02u.%03u] %s\r\n",
                            st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond, st.wMilliseconds, msg.c_str());
    std::lock_guard<std::mutex> lock(g_logMutex);
    HANDLE h = CreateFileW(g_logPath.c_str(), FILE_APPEND_DATA, FILE_SHARE_READ, nullptr, OPEN_ALWAYS, 0, nullptr);
    if (h == INVALID_HANDLE_VALUE) return;
    auto utf8 = WideToUtf8(line);
    DWORD w = 0;
    WriteFile(h, utf8.data(), (DWORD)utf8.size(), &w, nullptr);
    CloseHandle(h);
}

} // namespace util
