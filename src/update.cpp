// Self-update implementation.
#include "update.h"
#include "util.h"

namespace update {

static const wchar_t* kUA = L"SpoolerCleanup/" PSQC_VERSION_STR_W L" (Win32; update-check)";

namespace {
struct WinHttp {
    HINTERNET ses = nullptr, con = nullptr, req = nullptr;
    ~WinHttp() {
        if (req) WinHttpCloseHandle(req);
        if (con) WinHttpCloseHandle(con);
        if (ses) WinHttpCloseHandle(ses);
    }
};

bool HttpGet(const wchar_t* host, const wchar_t* path, std::string& body, DWORD& status, std::wstring& err) {
    WinHttp w;
    w.ses = WinHttpOpen(kUA, WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, nullptr, nullptr, 0);
    if (!w.ses) { err = util::LastErrorText(GetLastError()); return false; }
    DWORD opts = WINHTTP_FLAG_SECURE | WINHTTP_FLAG_ESCAPE_DISABLE;
    w.con = WinHttpConnect(w.ses, host, INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!w.con) { err = util::LastErrorText(GetLastError()); return false; }
    w.req = WinHttpOpenRequest(w.con, L"GET", path, nullptr, WINHTTP_NO_REFERER,
                               WINHTTP_DEFAULT_ACCEPT_TYPES, opts);
    if (!w.req) { err = util::LastErrorText(GetLastError()); return false; }
    WinHttpSetTimeouts(w.req, 10000, 10000, 15000, 30000);
    if (!WinHttpSendRequest(w.req, WINHTTP_NO_ADDITIONAL_HEADERS, 0, nullptr, 0, 0, 0)) {
        err = util::LastErrorText(GetLastError()); return false;
    }
    if (!WinHttpReceiveResponse(w.req, nullptr)) { err = util::LastErrorText(GetLastError()); return false; }
    DWORD sc = 0, sz = sizeof(sc);
    WinHttpQueryHeaders(w.req, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, nullptr, &sc, &sz, nullptr);
    status = sc;
    DWORD avail = 0;
    while (WinHttpQueryDataAvailable(w.req, &avail) && avail > 0) {
        std::vector<char> buf(avail);
        DWORD read = 0;
        if (!WinHttpReadData(w.req, buf.data(), avail, &read)) break;
        body.append(buf.data(), read);
    }
    return status == 200;
}
} // namespace

bool CheckForUpdate(ReleaseInfo& out, std::wstring& err) {
    out = ReleaseInfo{};
#ifdef PSQC_STORE_BUILD
    // Store packages are updated by the Store itself; in-app updater is disabled.
    err = L"disabled in Store build";
    return false;
#endif
    std::string body;
    DWORD status = 0;
    if (!HttpGet(L"api.github.com", L"/repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest", body, status, err)) {
        if (status == 404) err = L"no releases published";
        return false;
    }
    util::JsonValue root;
    if (!util::JsonParse(body, root)) { err = L"bad JSON"; return false; }
    auto* tag = root.Find(L"tag_name");
    if (!tag || tag->type != util::JsonValue::Str) { err = L"no tag_name"; return false; }
    out.tag = tag->str;
    if (auto* n = root.Find(L"name")) if (n->type == util::JsonValue::Str) out.name = n->str;
    if (auto* b = root.Find(L"body")) if (b->type == util::JsonValue::Str) out.notes = b->str;
    // pick asset: prefer *x64*.exe, else any .exe
    auto* assets = root.Find(L"assets");
    if (assets) {
        for (auto& kv : assets->obj) {
            const util::JsonValue& a = kv.second;
            auto* nm = a.Find(L"name");
            auto* url = a.Find(L"browser_download_url");
            if (!nm || !url || nm->type != util::JsonValue::Str || url->type != util::JsonValue::Str) continue;
            if (!util::EndsWithI(nm->str, L".exe")) continue;
            bool x64 = nm->str.find(L"x64") != std::wstring::npos;
            if (out.assetUrl.empty() || (x64 && out.assetName.find(L"x64") == std::wstring::npos)) {
                out.assetName = nm->str;
                out.assetUrl = url->str;
                if (auto* s = a.Find(L"size")) if (s->type == util::JsonValue::Num) out.size = (uint64_t)s->num;
            }
        }
    }
    out.valid = util::CompareVersions(out.tag, PSQC_VERSION_STR_W) > 0;
    return out.valid;
}

std::wstring UpdateTempPath() {
    return util::JoinPath(util::TempDir(), L"SpoolerCleanup-update.exe");
}

bool DownloadFile(const std::wstring& url, const std::wstring& dest,
                  const std::function<void(int)>& progress, const std::atomic<bool>* cancel) {
    // split url: https://host/path
    if (!util::StartsWithI(url, L"https://")) return false;
    auto rest = url.substr(8);
    auto slash = rest.find(L'/');
    std::wstring host = slash == std::wstring::npos ? std::wstring(rest) : std::wstring(rest.substr(0, slash));
    std::wstring path = slash == std::wstring::npos ? L"/" : std::wstring(rest.substr(slash));

    WinHttp w;
    w.ses = WinHttpOpen(kUA, WINHTTP_ACCESS_TYPE_AUTOMATIC_PROXY, nullptr, nullptr, 0);
    if (!w.ses) return false;
    w.con = WinHttpConnect(w.ses, host.c_str(), INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!w.con) return false;
    w.req = WinHttpOpenRequest(w.con, L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER,
                               WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
    if (!w.req) return false;
    WinHttpSetTimeouts(w.req, 10000, 10000, 20000, 120000);
    if (!WinHttpSendRequest(w.req, WINHTTP_NO_ADDITIONAL_HEADERS, 0, nullptr, 0, 0, 0)) return false;
    if (!WinHttpReceiveResponse(w.req, nullptr)) return false;

    uint64_t total = 0;
    DWORD sz = sizeof(total);
    WinHttpQueryHeaders(w.req, WINHTTP_QUERY_CONTENT_LENGTH | WINHTTP_QUERY_FLAG_NUMBER64, nullptr, &total, &sz, nullptr);

    HANDLE h = CreateFileW(dest.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    uint64_t done = 0;
    int lastPct = -1;
    bool ok = true;
    std::vector<char> buf(256 * 1024);
    while (true) {
        if (cancel && cancel->load()) { ok = false; break; }
        DWORD avail = 0;
        if (!WinHttpQueryDataAvailable(w.req, &avail)) { ok = false; break; }
        if (avail == 0) break;
        DWORD read = 0;
        if (!WinHttpReadData(w.req, buf.data(), (DWORD)buf.size(), &read) || read == 0) { ok = false; break; }
        DWORD written = 0;
        if (!WriteFile(h, buf.data(), read, &written, nullptr) || written != read) { ok = false; break; }
        done += read;
        if (total && progress) {
            int pct = (int)(done * 100 / total);
            if (pct != lastPct) { lastPct = pct; progress(pct); }
        }
    }
    CloseHandle(h);
    if (!ok) util::RemoveFile(dest);
    else if (progress) progress(100);
    return ok;
}

bool LaunchSelfReplace(const std::wstring& downloaded, const std::wstring& target) {
    std::wstring params = util::Fmt(L"/replace %lu \"%s\" \"%s\"", GetCurrentProcessId(), target.c_str(), downloaded.c_str());
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC;
    sei.lpVerb = L"runas";   // замена файлов в Program Files требует повышения
    sei.lpFile = downloaded.c_str();
    sei.lpParameters = params.c_str();
    sei.nShow = SW_HIDE;
    return ShellExecuteExW(&sei) != 0;
}

int RunSelfReplace(const std::wstring& pid, const std::wstring& target, const std::wstring& downloaded) {
    DWORD p = (DWORD)_wtoi(pid.c_str());
    if (p) {
        HANDLE hp = OpenProcess(SYNCHRONIZE, FALSE, p);
        if (hp) { WaitForSingleObject(hp, 20000); CloseHandle(hp); }
    }
    Sleep(500);
    std::wstring old = target + L".old";
    util::RemoveFile(old);
    bool moved = MoveFileExW(target.c_str(), old.c_str(), MOVEFILE_REPLACE_EXISTING) != 0;
    if (!moved) moved = !util::FileExists(target);
    bool ok = MoveFileExW(downloaded.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING) != 0;
    if (!ok) {
        // locked: откатываем старый бинарь на место — приложение не должно исчезнуть с диска
        DWORD err = GetLastError();
        if (moved) MoveFileExW(old.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING);
        // новый файл подложится при перезагрузке (операция требует прав администратора)
        MoveFileExW(downloaded.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_DELAY_UNTIL_REBOOT);
        SetLastError(err);
        return 3;
    }
    util::RemoveFile(old);
    SHELLEXECUTEINFOW sei{};
    sei.cbSize = sizeof(sei);
    sei.fMask = SEE_MASK_NOASYNC;
    sei.lpVerb = L"open";
    sei.lpFile = target.c_str();
    sei.lpParameters = L"/tray";
    sei.nShow = SW_SHOWNORMAL;
    ShellExecuteExW(&sei);
    return 0;
}

} // namespace update
