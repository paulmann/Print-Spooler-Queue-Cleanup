// Print spooler engine implementation. Safety model inherited from v4:
// only *.SPL / *.SHD inside %SystemRoot%\System32\spool\PRINTERS are ever touched.
#include "spooler.h"
#include "util.h"
#include "settings.h"

namespace spooler {

static bool IsSpoolFile(const wchar_t* name) {
    const wchar_t* dot = wcsrchr(name, L'.');
    if (!dot) return false;
    return util::IEquals(dot, L".spl") || util::IEquals(dot, L".shd");
}

int JobStatusStrId(DWORD s) {
    if (s & JOB_STATUS_DELETING) return STR_JOB_DELETING;
    if (s & JOB_STATUS_ERROR) return STR_JOB_ERROR;
    if (s & JOB_STATUS_PRINTING) return STR_JOB_PRINTING;
    if (s & JOB_STATUS_PAUSED) return STR_JOB_PAUSED;
    if (s & JOB_STATUS_OFFLINE) return STR_JOB_OFFLINE;
    if (s & JOB_STATUS_SPOOLING) return STR_JOB_SPOOLING;
    return STR_JOB_QUEUED;
}

DWORD ServiceState() {
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
    if (!scm) return 0;
    SC_HANDLE svc = OpenServiceW(scm, L"Spooler", SERVICE_QUERY_STATUS);
    if (!svc) { CloseServiceHandle(scm); return 0; }
    SERVICE_STATUS st{};
    DWORD state = 0;
    if (QueryServiceStatus(svc, &st)) state = st.dwCurrentState;
    CloseServiceHandle(svc);
    CloseServiceHandle(scm);
    return state;
}

bool EnumerateAll(std::vector<PrinterJobs>& out) {
    out.clear();
    DWORD needed = 0, count = 0;
    EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, nullptr, 2, nullptr, 0, &needed, &count);
    if (count == 0 && needed == 0) return true;
    std::vector<BYTE> buf(needed + 8);
    if (!EnumPrintersW(PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS, nullptr, 2, buf.data(), needed, &needed, &count))
        return false;
    auto* info = (PRINTER_INFO_2W*)buf.data();
    for (DWORD i = 0; i < count; ++i) {
        if (!info[i].pPrinterName) continue;
        PrinterJobs pj;
        pj.name = info[i].pPrinterName;
        pj.printerStatus = info[i].Status;
        pj.paused = (info[i].Attributes & PRINTER_ATTRIBUTE_WORK_OFFLINE) != 0 ||
                    (info[i].Status & PRINTER_STATUS_PAUSED) != 0;
        HANDLE h = nullptr;
        if (OpenPrinterW(info[i].pPrinterName, &h, nullptr)) {
            DWORD jneed = 0, jcount = 0;
            EnumJobsW(h, 0, 256, 2, nullptr, 0, &jneed, &jcount);
            if (jneed > 0) {
                std::vector<BYTE> jb(jneed + 8);
                if (EnumJobsW(h, 0, 256, 2, jb.data(), jneed, &jneed, &jcount)) {
                    auto* jobs = (JOB_INFO_2W*)jb.data();
                    for (DWORD j = 0; j < jcount; ++j) {
                        JobInfo ji;
                        ji.jobId = jobs[j].JobId;
                        ji.printer = pj.name;
                        ji.document = jobs[j].pDocument ? jobs[j].pDocument : L"";
                        ji.owner = jobs[j].pUserName ? jobs[j].pUserName : L"";
                        ji.size = jobs[j].Size;
                        ji.status = jobs[j].Status;
                        ji.pagesPrinted = jobs[j].PagesPrinted;
                        ji.totalPages = jobs[j].TotalPages;
                        if (jobs[j].Submitted.wYear) ji.submitted = jobs[j].Submitted;
                        else GetLocalTime(&ji.submitted);
                        pj.jobs.push_back(std::move(ji));
                    }
                }
            }
            ClosePrinter(h);
        }
        out.push_back(std::move(pj));
    }
    return true;
}

bool DeleteJob(const std::wstring& printer, DWORD jobId, DWORD& err) {
    HANDLE h = nullptr;
    if (!OpenPrinterW((LPWSTR)printer.c_str(), &h, nullptr)) { err = GetLastError(); return false; }
    BOOL ok = SetJobW(h, jobId, 0, nullptr, JOB_CONTROL_DELETE);
    if (!ok) err = GetLastError();
    ClosePrinter(h);
    return ok != 0;
}

bool PausePrinter(const std::wstring& printer, bool pause, DWORD& err) {
    HANDLE h = nullptr;
    if (!OpenPrinterW((LPWSTR)printer.c_str(), &h, nullptr)) { err = GetLastError(); return false; }
    BOOL ok = SetPrinterW(h, 0, nullptr, pause ? PRINTER_CONTROL_PAUSE : PRINTER_CONTROL_RESUME);
    if (!ok) err = GetLastError();
    ClosePrinter(h);
    return ok != 0;
}

bool PauseAll(bool pause) {
    std::vector<PrinterJobs> all;
    if (!EnumerateAll(all)) return false;
    bool ok = true;
    for (auto& p : all) {
        DWORD err = 0;
        if (p.jobs.empty() && !pause) continue;
        if (!PausePrinter(p.name, pause, err)) ok = false;
    }
    return ok;
}

bool AnyPaused(std::vector<PrinterJobs>& snapshot) {
    for (auto& p : snapshot)
        if (p.paused) return true;
    return false;
}

// ---- service control --------------------------------------------------------
static bool WaitState(SC_HANDLE svc, DWORD want, DWORD avoid, int timeoutMs) {
    SERVICE_STATUS st{};
    for (int t = 0; t < timeoutMs; t += 200) {
        if (!QueryServiceStatus(svc, &st)) return false;
        if (st.dwCurrentState == want) return true;
        if (avoid && st.dwCurrentState == avoid) return false;
        Sleep(200);
    }
    return QueryServiceStatus(svc, &st) && st.dwCurrentState == want;
}

bool StopService(std::wstring& err) {
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
    if (!scm) { err = util::LastErrorText(GetLastError()); return false; }
    SC_HANDLE svc = OpenServiceW(scm, L"Spooler", SERVICE_STOP | SERVICE_QUERY_STATUS);
    if (!svc) { err = util::LastErrorText(GetLastError()); CloseServiceHandle(scm); return false; }
    SERVICE_STATUS st{};
    bool ok = ControlService(svc, SERVICE_CONTROL_STOP, &st) != 0;
    if (!ok && GetLastError() == ERROR_SERVICE_NOT_ACTIVE) ok = true;   // уже остановлена — цель достигнута
    if (ok) ok = WaitState(svc, SERVICE_STOPPED, 0, 15000);
    if (!ok) err = util::LastErrorText(GetLastError());
    CloseServiceHandle(svc);
    CloseServiceHandle(scm);
    return ok;
}

bool StartService_(std::wstring& err) {
    SC_HANDLE scm = OpenSCManagerW(nullptr, nullptr, SC_MANAGER_CONNECT);
    if (!scm) { err = util::LastErrorText(GetLastError()); return false; }
    SC_HANDLE svc = OpenServiceW(scm, L"Spooler", SERVICE_START | SERVICE_QUERY_STATUS);
    if (!svc) { err = util::LastErrorText(GetLastError()); CloseServiceHandle(scm); return false; }
    bool ok = ::StartServiceW(svc, 0, nullptr) != 0;
    if (!ok && GetLastError() == ERROR_SERVICE_ALREADY_RUNNING) ok = true;
    if (ok) ok = WaitState(svc, SERVICE_RUNNING, 0, 15000);
    if (!ok) err = util::LastErrorText(GetLastError());
    CloseServiceHandle(svc);
    CloseServiceHandle(scm);
    return ok;
}

bool RestartService(std::wstring& err) {
    DWORD state = ServiceState();
    if (state == SERVICE_RUNNING || state == SERVICE_START_PENDING || state == SERVICE_STOP_PENDING) {
        if (!StopService(err)) return false;
    }
    return StartService_(err);
}

// ---- backup / restore / purge ----------------------------------------------
bool BackupSpoolFiles(std::wstring& outDir, int& outCount) {
    outCount = 0;
    auto src = util::SpoolDir();
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(util::JoinPath(src, L"*").c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return true;   // nothing to back up
    std::vector<std::wstring> names;
    do {
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && IsSpoolFile(fd.cFileName))
            names.emplace_back(fd.cFileName);
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    if (names.empty()) return true;
    auto dst = util::JoinPath(util::BackupRootDir(true), util::NowStamp());
    if (!util::EnsureDir(dst)) return false;
    for (auto& n : names) {
        if (CopyFileW(util::JoinPath(src, n).c_str(), util::JoinPath(dst, n).c_str(), FALSE))
            ++outCount;
    }
    outDir = dst;
    return outCount > 0;
}

std::vector<BackupInfo> ListBackups() {
    std::vector<BackupInfo> out;
    auto root = util::BackupRootDir(false);
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(util::JoinPath(root, L"*").c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return out;
    do {
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) ||
            wcscmp(fd.cFileName, L".") == 0 || wcscmp(fd.cFileName, L"..") == 0) continue;
        BackupInfo bi;
        bi.path = util::JoinPath(root, fd.cFileName);
        FILETIME lt{};
        FileTimeToLocalFileTime(&fd.ftCreationTime, &lt);
        FileTimeToSystemTime(&lt, &bi.when);
        WIN32_FIND_DATAW f2{};
        HANDLE h2 = FindFirstFileW(util::JoinPath(bi.path, L"*").c_str(), &f2);
        if (h2 != INVALID_HANDLE_VALUE) {
            do {
                if (f2.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
                ++bi.fileCount;
                bi.totalSize += ((uint64_t)f2.nFileSizeHigh << 32) | f2.nFileSizeLow;
            } while (FindNextFileW(h2, &f2));
            FindClose(h2);
        }
        out.push_back(std::move(bi));
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    std::sort(out.begin(), out.end(), [](const BackupInfo& a, const BackupInfo& b) { return a.path > b.path; });
    return out;
}

int SpoolFileCount() {
    int n = 0;
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(util::JoinPath(util::SpoolDir(), L"*").c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return 0;
    do {
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && IsSpoolFile(fd.cFileName)) ++n;
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    return n;
}

static int CountAllJobs() {
    std::vector<PrinterJobs> all;
    if (!EnumerateAll(all)) return -1;
    int n = 0;
    for (auto& p : all) n += (int)p.jobs.size();
    return n;
}

static void SweepSpoolDir() {
    auto dir = util::SpoolDir();
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(util::JoinPath(dir, L"*").c_str(), &fd);
    if (h == INVALID_HANDLE_VALUE) return;
    std::vector<std::wstring> names;
    do {
        if (!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) && IsSpoolFile(fd.cFileName))
            names.emplace_back(fd.cFileName);
    } while (FindNextFileW(h, &fd));
    FindClose(h);
    for (auto& n : names) {
        auto full = util::JoinPath(dir, n);
        for (int attempt = 0; attempt < 3; ++attempt) {
            SetFileAttributesW(full.c_str(), FILE_ATTRIBUTE_NORMAL);
            if (DeleteFileW(full.c_str())) break;
            Sleep(200);
        }
    }
}

bool RestoreSpoolFiles(const std::wstring& backupDir, PurgeResult& res) {
    if (!util::DirExists(backupDir)) { res.error = L"backup dir missing"; res.errCode = ERROR_PATH_NOT_FOUND; return false; }
    std::wstring err;
    if (!StopService(err)) { res.error = err; res.errCode = GetLastError(); return false; }
    auto dst = util::SpoolDir();
    WIN32_FIND_DATAW fd{};
    HANDLE h = FindFirstFileW(util::JoinPath(backupDir, L"*").c_str(), &fd);
    int copied = 0;
    if (h != INVALID_HANDLE_VALUE) {
        do {
            if (fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) continue;
            if (!IsSpoolFile(fd.cFileName)) continue;
            if (CopyFileW(util::JoinPath(backupDir, fd.cFileName).c_str(),
                          util::JoinPath(dst, fd.cFileName).c_str(), FALSE))
                ++copied;
        } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
    res.restoredFiles = copied;
    bool started = StartService_(err);
    if (!started) { res.error = err; res.errCode = GetLastError(); return false; }
    res.ok = copied > 0;
    return res.ok;
}

void PurgeAll(bool backup, PurgeResult& res) {
    res = PurgeResult{};
    res.removedBefore = CountAllJobs();
    if (res.removedBefore < 0) { res.error = L"enumerate failed"; res.errCode = GetLastError(); return; }

    if (backup && res.removedBefore > 0) {
        int n = 0;
        std::wstring dir;
        if (BackupSpoolFiles(dir, n)) res.backupPath = dir;
    }

    // 1) API purge per printer
    std::vector<PrinterJobs> all;
    if (EnumerateAll(all)) {
        for (auto& p : all) {
            if (p.jobs.empty()) continue;
            HANDLE h = nullptr;
            if (OpenPrinterW((LPWSTR)p.name.c_str(), &h, nullptr)) {
                SetPrinterW(h, 0, nullptr, PRINTER_CONTROL_PURGE);
                ClosePrinter(h);
            }
        }
    }
    Sleep(300);
    res.removedAfter = CountAllJobs();

    // 2) stubborn leftovers: service restart + guarded file sweep
    if (res.removedAfter > 0 || SpoolFileCount() > 0) {
        std::wstring err;
        if (StopService(err)) {
            SweepSpoolDir();
            res.usedServiceRestart = StartService_(err);
            if (!res.usedServiceRestart) { res.error = err; res.errCode = GetLastError(); return; }
            Sleep(300);
            res.removedAfter = CountAllJobs();
        } else {
            res.error = err;
            res.errCode = GetLastError();
            return;
        }
    }
    res.ok = res.removedAfter == 0;
    if (!res.ok && res.error.empty()) { res.error = L"jobs remained after purge"; res.errCode = ERROR_ACCESS_DENIED; }
}

} // namespace spooler
