#pragma once
// Single source of truth for product identity & versioning.
// Bump PSQC_VERSION_* here only; everything (rc, About, MSIX, update check) derives from it.
#define PSQC_PRODUCT_NAME      L"Print Spooler Queue Cleanup"
#define PSQC_EXE_NAME          L"SpoolerCleanup.exe"
#define PSQC_COMPANY           L"Mikhail Deynekin"
#define PSQC_AUTHOR_EMAIL      L"icons@Deynekin.com"
#define PSQC_HOMEPAGE          "https://Deynekin.com"
#define PSQC_REPO_URL          "https://github.com/paulmann/Print-Spooler-Queue-Cleanup"
#define PSQC_WIKI_URL          "https://github.com/paulmann/Print-Spooler-Queue-Cleanup/wiki"
#define PSQC_LICENSE_URL       "https://github.com/paulmann/Print-Spooler-Queue-Cleanup/blob/main/LICENSE"
#define PSQC_LICENSE_NAME      L"PolyForm Noncommercial License 1.0.0"
#define PSQC_ISSUES_URL        "https://github.com/paulmann/Print-Spooler-Queue-Cleanup/issues"
#define PSQC_RELEASES_API      L"https://api.github.com/repos/paulmann/Print-Spooler-Queue-Cleanup/releases/latest"

#define PSQC_VERSION_MAJOR 5
#define PSQC_VERSION_MINOR 1
#define PSQC_VERSION_PATCH 0
#define PSQC_VERSION_BUILD 0

#define PSQC_VERSION_STR_A "5.1.0"
#define PSQC_VERSION_STR_W L"5.1.0"
#define PSQC_VERSION_QUAD  5,1,0,0

// Settings / instance identity
#define PSQC_SETTINGS_DIR      L"SpoolerCleanup"
#define PSQC_MUTEX_NAME        L"Local\\SpoolerCleanup_SingleInstance"
#define PSQC_WINDOW_CLASS      L"SpoolerCleanupMainClass"
#define PSQC_WIZARD_CLASS      L"SpoolerCleanupWizardClass"
#define PSQC_DIALOG_CLASS      L"SpoolerCleanupDialogClass"
#define PSQC_ABOUT_CLASS       L"SpoolerCleanupAboutClass"
