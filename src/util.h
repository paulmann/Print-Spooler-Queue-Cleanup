#pragma once
// Utility layer: strings, paths, registry, files, system, DPI, logging, clipboard.
#include "app.h"
#include "pure.h"   // json / semver: platform-independent, unit-tested

namespace util {

// ---- strings ---------------------------------------------------------------
std::wstring Utf8ToWide(std::string_view utf8);
std::string  WideToUtf8(std::wstring_view wide);
std::wstring Fmt(const wchar_t* fmt, ...);
std::wstring Trim(std::wstring_view s);
std::wstring Lower(std::wstring_view s);
bool         IEquals(std::wstring_view a, std::wstring_view b);
bool         StartsWithI(std::wstring_view s, std::wstring_view prefix);
bool         EndsWithI(std::wstring_view s, std::wstring_view suffix);
std::vector<std::wstring> Split(std::wstring_view s, wchar_t sep);
std::wstring FormatBytes(uint64_t bytes);            // invariant "1.25 MB"
std::wstring FormatInt(uint64_t v);                  // no thousand sep (stable for CLI)

// ---- paths -----------------------------------------------------------------
std::wstring ExePath();
std::wstring ExeDir();
std::wstring AppDataDir(bool create);                // %APPDATA%\SpoolerCleanup
std::wstring LocalDataDir(bool create);              // %LOCALAPPDATA%\SpoolerCleanup
std::wstring LogsDir(bool create);
std::wstring BackupRootDir(bool create);
std::wstring SettingsFilePath();                     // respects portable mode
std::wstring JoinPath(std::wstring_view dir, std::wstring_view file);
std::wstring SpoolDir();                              // %SystemRoot%\System32\spool\PRINTERS
std::wstring TempDir();

// ---- files -----------------------------------------------------------------
bool FileExists(const std::wstring& path);
bool DirExists(const std::wstring& path);
bool EnsureDir(const std::wstring& path);
bool ReadFileAll(const std::wstring& path, std::string& out);
bool WriteFileAll(const std::wstring& path, std::string_view data);
bool RemoveFile(const std::wstring& path);
bool RemoveDirTree(const std::wstring& path);
uint64_t FileSize(const std::wstring& path);

// ---- registry --------------------------------------------------------------
bool RegGetDword(HKEY root, const wchar_t* subkey, const wchar_t* name, DWORD& out);
bool RegSetDword(HKEY root, const wchar_t* subkey, const wchar_t* name, DWORD value);
bool RegGetString(HKEY root, const wchar_t* subkey, const wchar_t* name, std::wstring& out);
bool RegSetString(HKEY root, const wchar_t* subkey, const wchar_t* name, const wchar_t* value);
bool RegDeleteValue(HKEY root, const wchar_t* subkey, const wchar_t* name);
bool RegDeleteTree(HKEY root, const wchar_t* subkey);

// ---- system ----------------------------------------------------------------
bool IsElevated();
bool IsWineRuntime();              // classic ntdll!wine_get_version probe
DWORD WindowsBuild();
void OpenUrl(const char* url);
bool ShellRun(const wchar_t* verb, const wchar_t* file, const wchar_t* params, bool wait, DWORD* exitCode);
bool RunElevatedAndWait(const wchar_t* file, const wchar_t* params, DWORD* exitCode);
std::wstring DefaultPrinterName();
void OpenPrinterQueueUI(const std::wstring& printerName);   // falls back to printers folder
void OpenPrintersFolder();                                   // classic control panel printers
void OpenPrintersSettingsPage();                             // ms-settings:printers
void OpenFolder(const std::wstring& path);
std::wstring LastErrorText(DWORD err);
std::wstring OsFriendlyName();                               // "Windows 11 Pro (build 26200)"
std::wstring CurrentUserSid();

// ---- time ------------------------------------------------------------------
std::wstring FormatSystemTimeUser(const SYSTEMTIME& st);     // user-locale date + time
std::wstring NowStamp();                                     // yyyyMMdd-HHmmss (file names)
uint64_t UnixSeconds();

// ---- dpi -------------------------------------------------------------------
int  DpiOf(HWND hwnd);
int  Scale(HWND hwnd, int dip);                              // dip * dpi / 96

// ---- misc ------------------------------------------------------------------
void ApplyRtlLayout(HWND root);          // mirror window + children for RTL languages
bool QaShot(HWND hwnd, const wchar_t* name);  // self-capture to %PSQC_SHOTDIR% (QA hook)
bool CopyToClipboard(const std::wstring& text);
HICON LoadAppIcon(int cx, int cy);
std::wstring LoadStringRes(UINT id);                          // rc strings (fallback only)
void CenterWindow(HWND hwnd, HWND owner);
RECT WorkAreaRect();

// ---- logging ---------------------------------------------------------------
enum LogLevel : int { LOG_OFF = 0, LOG_ERR = 1, LOG_ALL = 2 };
void LogInit();
void LogClose();
void Log(int level, const wchar_t* fmt, ...);
void SetLogLevel(int lvl);
int  GetLogLevel();
#define LOGE(...) ::util::Log(::util::LOG_ERR, __VA_ARGS__)
#define LOGA(...) ::util::Log(::util::LOG_ALL, __VA_ARGS__)

} // namespace util
