#pragma once
// Print spooler engine: enumeration, deletion, purge-with-backup, service control.
#include "app.h"
#include "i18n.h"

namespace spooler {

struct JobInfo {
    DWORD jobId = 0;
    std::wstring printer;
    std::wstring document;
    std::wstring owner;
    uint64_t size = 0;
    DWORD status = 0;
    SYSTEMTIME submitted{};
    DWORD pagesPrinted = 0;
    DWORD totalPages = 0;
};

struct PrinterJobs {
    std::wstring name;
    DWORD printerStatus = 0;
    bool paused = false;
    std::vector<JobInfo> jobs;
};

struct PurgeResult {
    bool ok = false;
    int removedBefore = 0;      // jobs seen before purge
    int removedAfter = 0;       // jobs still present after purge
    int restoredFiles = 0;
    std::wstring backupPath;
    std::wstring error;
    DWORD errCode = 0;
    bool usedServiceRestart = false;
};

struct BackupInfo {
    std::wstring path;
    SYSTEMTIME when{};
    int fileCount = 0;
    uint64_t totalSize = 0;
};

bool EnumerateAll(std::vector<PrinterJobs>& out);
DWORD ServiceState();                       // SERVICE_* or 0
int   JobStatusStrId(DWORD status);         // StrId for status column
bool  DeleteJob(const std::wstring& printer, DWORD jobId, DWORD& err);
bool  PausePrinter(const std::wstring& printer, bool pause, DWORD& err);
bool  PauseAll(bool pause);
bool  AnyPaused(std::vector<PrinterJobs>& snapshot);
bool  BackupSpoolFiles(std::wstring& outDir, int& outCount);
bool  RestoreSpoolFiles(const std::wstring& backupDir, PurgeResult& res);
void  PurgeAll(bool backup, PurgeResult& res);          // API purge, fallback: service restart + file sweep
bool  RestartService(std::wstring& err);
bool  StopService(std::wstring& err);
bool  StartService_(std::wstring& err);
std::vector<BackupInfo> ListBackups();
int   SpoolFileCount();

} // namespace spooler
