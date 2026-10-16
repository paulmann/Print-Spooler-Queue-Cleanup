#pragma once
// Self-update: GitHub Releases check (WinHTTP), download, in-place replace.
#include "app.h"

namespace update {

struct ReleaseInfo {
    std::wstring tag;            // "v5.0.1"
    std::wstring name;
    std::wstring assetUrl;
    std::wstring assetName;
    uint64_t size = 0;
    std::wstring notes;
    bool valid = false;
};

// true = newer release found; err filled on transport failure.
bool CheckForUpdate(ReleaseInfo& out, std::wstring& err);
bool DownloadFile(const std::wstring& url, const std::wstring& dest,
                  const std::function<void(int percent)>& progress,
                  const std::atomic<bool>* cancel);
// Spawn "<exe> /replace <pid> <target> <downloaded>" detached; caller should exit.
bool LaunchSelfReplace(const std::wstring& downloaded, const std::wstring& target);
// Runs in the spawned child: waits for parent, swaps files, relaunches. Returns process exit code.
int  RunSelfReplace(const std::wstring& pid, const std::wstring& target, const std::wstring& downloaded);
std::wstring UpdateTempPath();

} // namespace update
