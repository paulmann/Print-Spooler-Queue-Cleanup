#pragma once
// Internationalization: 30 languages, system-language default, RTL support.
// Tables live in i18n_tables_a/b/c.cpp. ORDER OF IDS IS FROZEN — every table
// must contain exactly STR_COUNT entries in this order.
#include "app.h"

#include "i18n_ids.h"

struct LangInfo {
    const wchar_t* tag;      // "en", "ru", "zh-Hans", ...
    const wchar_t* native;   // native name, shown in pickers
    bool rtl;
};

const LangInfo* Langs();          // array of 30
int  LangCount();                 // 30
int  CurrentLang();               // index into Langs()
void SetLanguage(const wchar_t* tag);   // "" or nullptr = follow system
std::wstring DetectSystemLangTag();
bool CurrentLangIsRtl();
const wchar_t* Str(int id);       // localized string for current language
template <typename... Args>
std::wstring StrF(int id, Args... args) {   // localized + printf-style
    extern std::wstring Fmt(const wchar_t*, ...);
    return Fmt(Str(id), args...);
}

