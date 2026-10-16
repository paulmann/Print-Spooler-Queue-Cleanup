#pragma once
// Frozen string-id enum + translation table externs. NO Windows dependencies:
// this header (and the tables) compile on any platform for integrity tests.
#include <cstddef>

enum StrId : int {
    STR_APP_TITLE = 0,      // 0
    STR_APP_TAGLINE,        // 1
    STR_TB_REFRESH,         // 2
    STR_TB_CLEAR_SEL,       // 3
    STR_TB_CLEAR_ALL,       // 4
    STR_TB_PAUSE,           // 5
    STR_TB_RESUME,          // 6
    STR_TB_RESTART,         // 7
    STR_TB_SETTINGS,        // 8
    STR_TB_ABOUT,           // 9
    STR_COL_PRINTER,        // 10
    STR_COL_DOCUMENT,       // 11
    STR_COL_OWNER,          // 12
    STR_COL_SIZE,           // 13
    STR_COL_STATUS,         // 14
    STR_COL_SUBMITTED,      // 15
    STR_EMPTY_TITLE,        // 16
    STR_EMPTY_HINT,         // 17
    STR_ST_JOBS,            // 18  %s
    STR_ST_PRINTERS,        // 19  %s
    STR_ST_UPDATED,         // 20  %s
    STR_ST_SERVICE,         // 21  %s
    STR_ST_ADMIN,           // 22
    STR_ST_USER,            // 23
    STR_SVC_RUNNING,        // 24
    STR_SVC_STOPPED,        // 25
    STR_SVC_PENDING,        // 26
    STR_JOB_PRINTING,       // 27
    STR_JOB_SPOOLING,       // 28
    STR_JOB_PAUSED,         // 29
    STR_JOB_ERROR,          // 30
    STR_JOB_DELETING,       // 31
    STR_JOB_OFFLINE,        // 32
    STR_JOB_QUEUED,         // 33
    STR_MSG_CONFIRM_ALL,    // 34
    STR_MSG_CONFIRM_SEL,    // 35  %s
    STR_MSG_CLEAR_OK,       // 36  %s
    STR_MSG_CLEAR_FAIL,     // 37  %s
    STR_MSG_NEED_ADMIN,     // 38
    STR_MSG_ELEVATE_Q,      // 39
    STR_MSG_RESTART_OK,     // 40
    STR_MSG_STOP_FAIL,      // 41  %s
    STR_MSG_BACKUP_SAVED,   // 42  %s
    STR_MSG_RESTORE_OK,     // 43  %s
    STR_MSG_RESTORE_FAIL,   // 44  %s
    STR_MSG_UPDATE_NONE,    // 45  %s
    STR_MSG_UPDATE_AVAIL,   // 46  %s
    STR_MSG_UPDATE_FAIL,    // 47  %s
    STR_MSG_DIAG_COPIED,    // 48
    STR_DL_SETTINGS,        // 49
    STR_SET_APPEARANCE,     // 50
    STR_SET_LANGUAGE,       // 51
    STR_SET_THEME,          // 52
    STR_THEME_SYSTEM,       // 53
    STR_THEME_LIGHT,        // 54
    STR_THEME_DARK,         // 55
    STR_SET_BEHAVIOR,       // 56
    STR_SET_TRAY,           // 57
    STR_SET_MIN_TRAY,       // 58
    STR_SET_START_MIN,      // 59
    STR_SET_AUTOSTART,      // 60
    STR_SET_NOTIFY,         // 61
    STR_SET_CONFIRM,        // 62
    STR_SET_AUTOUPD,        // 63
    STR_SET_ELEVATE,        // 64
    STR_SET_BACKUP,         // 65
    STR_SET_REFRESH,        // 66
    STR_SET_LOGGING,        // 67
    STR_LOG_OFF,            // 68
    STR_LOG_ERRORS,         // 69
    STR_LOG_ALL,            // 70
    STR_SET_DATA,           // 71
    STR_BTN_OPEN_LOGS,      // 72
    STR_BTN_OPEN_BACKUPS,   // 73
    STR_BTN_RESTORE,        // 74
    STR_BTN_UNINSTALL,      // 75
    STR_RESTORE_TITLE,      // 76
    STR_RESTORE_HINT,       // 77
    STR_BTN_RESTORE_SEL,    // 78
    STR_WIZ_TITLE,          // 79
    STR_WIZ_WELCOME,        // 80
    STR_WIZ_WELCOME_TEXT,   // 81  %s
    STR_WIZ_RUN_ONCE,       // 82
    STR_WIZ_INSTALL,        // 83
    STR_WIZ_SCOPE,          // 84
    STR_WIZ_SCOPE_ALL,      // 85
    STR_WIZ_SCOPE_CUR,      // 86
    STR_WIZ_FOLDER,         // 87
    STR_WIZ_BROWSE,         // 88
    STR_WIZ_OPTIONS,        // 89
    STR_WIZ_OPT_DESKTOP,    // 90
    STR_WIZ_OPT_STARTMENU,  // 91
    STR_WIZ_OPT_AUTOSTART,  // 92
    STR_WIZ_OPT_TRAY,       // 93
    STR_WIZ_INSTALLING,     // 94
    STR_WIZ_DONE,           // 95
    STR_WIZ_DONE_TEXT,      // 96  %s
    STR_WIZ_LAUNCH,         // 97
    STR_WIZ_NEXT,           // 98
    STR_WIZ_BACK,           // 99
    STR_WIZ_CANCEL,         // 100
    STR_WIZ_ERR_COPY,       // 101 %s
    STR_WIZ_ERR_ADMIN,      // 102
    STR_ABOUT_TITLE,        // 103
    STR_ABOUT_VERSION,      // 104 %s
    STR_ABOUT_AUTHOR,       // 105
    STR_BTN_CHECK_UPD,      // 106
    STR_BTN_COPY_DIAG,      // 107
    STR_BTN_CLOSE,          // 108
    STR_TRAY_OPEN,          // 109
    STR_TRAY_JOBS,          // 110 %s
    STR_TRAY_QUEUE,         // 111
    STR_TRAY_PRINTERS,      // 112
    STR_TRAY_CLEAR,         // 113
    STR_TRAY_PAUSE,         // 114
    STR_TRAY_RESUME,         // 115
    STR_TRAY_RESTART,       // 116
    STR_TRAY_THEME,         // 117
    STR_TRAY_AUTOSTART,     // 118
    STR_TRAY_UPDATE,        // 119
    STR_TRAY_ABOUT,         // 120
    STR_TRAY_EXIT,          // 121
    STR_BTN_YES,            // 122
    STR_BTN_NO,             // 123
    STR_BTN_OK,             // 124
    STR_BTN_CANCEL,         // 125
    STR_BTN_RETRY,          // 126
    STR_BTN_LATER,          // 127
    STR_UNINSTALL_Q,        // 128 %s
    STR_UNINSTALL_DONE,     // 129
    STR_PORTABLE_BADGE,     // 130
    STR_LANG_NATIVE,        // 131 (native language name)
    STR_UPD_CHECKING,       // 132
    STR_UPD_DOWNLOADING,    // 133 %s
    STR_UPD_APPLYING,       // 134
    STR_LINK_SITE,          // 135
    STR_LINK_REPO,          // 136
    STR_LINK_LICENSE,       // 137
    STR_LINK_WIKI,          // 138
    STR_COPYRIGHT,          // 139
    STR_SET_PORTABLE_NOTE,  // 140
    STR_MSG_SEL_NONE,       // 141
    STR_BALLOON_CLEAR_OK,   // 142
    STR_BALLOON_UPDATE,     // 143 %s
    STR_WIZ_SCOPE_NOTE,     // 144
    STR_COUNT,              // 145
};

extern const wchar_t* const kStr_en[STR_COUNT];
extern const wchar_t* const kStr_ru[STR_COUNT];
extern const wchar_t* const kStr_es[STR_COUNT];
extern const wchar_t* const kStr_zhHans[STR_COUNT];
extern const wchar_t* const kStr_zhHant[STR_COUNT];
extern const wchar_t* const kStr_ar[STR_COUNT];
extern const wchar_t* const kStr_hi[STR_COUNT];
extern const wchar_t* const kStr_pt[STR_COUNT];
extern const wchar_t* const kStr_fr[STR_COUNT];
extern const wchar_t* const kStr_de[STR_COUNT];
extern const wchar_t* const kStr_ja[STR_COUNT];
extern const wchar_t* const kStr_it[STR_COUNT];
extern const wchar_t* const kStr_tr[STR_COUNT];
extern const wchar_t* const kStr_vi[STR_COUNT];
extern const wchar_t* const kStr_id[STR_COUNT];
extern const wchar_t* const kStr_th[STR_COUNT];
extern const wchar_t* const kStr_nl[STR_COUNT];
extern const wchar_t* const kStr_pl[STR_COUNT];
extern const wchar_t* const kStr_uk[STR_COUNT];
extern const wchar_t* const kStr_ro[STR_COUNT];
extern const wchar_t* const kStr_cs[STR_COUNT];
extern const wchar_t* const kStr_sv[STR_COUNT];
extern const wchar_t* const kStr_hu[STR_COUNT];
extern const wchar_t* const kStr_el[STR_COUNT];
extern const wchar_t* const kStr_he[STR_COUNT];
extern const wchar_t* const kStr_fa[STR_COUNT];
extern const wchar_t* const kStr_da[STR_COUNT];
extern const wchar_t* const kStr_fi[STR_COUNT];
extern const wchar_t* const kStr_nb[STR_COUNT];
extern const wchar_t* const kStr_ko[STR_COUNT];
