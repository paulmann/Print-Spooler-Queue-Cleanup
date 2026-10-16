// i18n dispatcher: language table registry, system detection, lookup.
#include "i18n.h"
#include "util.h"

static const LangInfo kLangs[] = {
    { L"en",      L"English",            false },
    { L"ru",      L"Русский",            false },
    { L"es",      L"Español",            false },
    { L"zh-Hans", L"简体中文",            false },
    { L"zh-Hant", L"繁體中文",            false },
    { L"ar",      L"العربية",            true  },
    { L"hi",      L"हिन्दी",              false },
    { L"pt",      L"Português",          false },
    { L"fr",      L"Français",           false },
    { L"de",      L"Deutsch",            false },
    { L"ja",      L"日本語",              false },
    { L"ko",      L"한국어",              false },
    { L"it",      L"Italiano",           false },
    { L"tr",      L"Türkçe",             false },
    { L"vi",      L"Tiếng Việt",         false },
    { L"id",      L"Bahasa Indonesia",   false },
    { L"th",      L"ไทย",                false },
    { L"nl",      L"Nederlands",         false },
    { L"pl",      L"Polski",             false },
    { L"uk",      L"Українська",         false },
    { L"ro",      L"Română",             false },
    { L"cs",      L"Čeština",            false },
    { L"sv",      L"Svenska",            false },
    { L"hu",      L"Magyar",             false },
    { L"el",      L"Ελληνικά",           false },
    { L"he",      L"עברית",              true  },
    { L"fa",      L"فارسی",              true  },
    { L"da",      L"Dansk",              false },
    { L"fi",      L"Suomi",              false },
    { L"nb",      L"Norsk (bokmål)",     false },
};

struct TableRef { const wchar_t* tag; const wchar_t* const* table; };
static const TableRef kTables[] = {
    { L"en", kStr_en }, { L"ru", kStr_ru }, { L"es", kStr_es }, { L"zh-Hans", kStr_zhHans },
    { L"zh-Hant", kStr_zhHant }, { L"ar", kStr_ar }, { L"hi", kStr_hi }, { L"pt", kStr_pt },
    { L"fr", kStr_fr }, { L"de", kStr_de }, { L"ja", kStr_ja }, { L"ko", kStr_ko },
    { L"it", kStr_it }, { L"tr", kStr_tr }, { L"vi", kStr_vi }, { L"id", kStr_id },
    { L"th", kStr_th }, { L"nl", kStr_nl }, { L"pl", kStr_pl }, { L"uk", kStr_uk },
    { L"ro", kStr_ro }, { L"cs", kStr_cs }, { L"sv", kStr_sv }, { L"hu", kStr_hu },
    { L"el", kStr_el }, { L"he", kStr_he }, { L"fa", kStr_fa }, { L"da", kStr_da },
    { L"fi", kStr_fi }, { L"nb", kStr_nb },
};

const LangInfo* Langs() { return kLangs; }
int LangCount() { return (int)(sizeof(kLangs) / sizeof(kLangs[0])); }

std::wstring DetectSystemLangTag() {
    LANGID lid = GetUserDefaultUILanguage();
    switch (PRIMARYLANGID(lid)) {
        case LANG_ENGLISH:    return L"en";
        case LANG_RUSSIAN:    return L"ru";
        case LANG_SPANISH:    return L"es";
        case LANG_CHINESE: {
            switch (SUBLANGID(lid)) {
                case SUBLANG_CHINESE_TRADITIONAL:
                case SUBLANG_CHINESE_HONGKONG:
                case SUBLANG_CHINESE_MACAU:
                    return L"zh-Hant";
                default:
                    return L"zh-Hans";
            }
        }
        case LANG_ARABIC:     return L"ar";
        case LANG_HINDI:      return L"hi";
        case LANG_PORTUGUESE: return L"pt";
        case LANG_FRENCH:     return L"fr";
        case LANG_GERMAN:     return L"de";
        case LANG_JAPANESE:   return L"ja";
        case LANG_KOREAN:     return L"ko";
        case LANG_ITALIAN:    return L"it";
        case LANG_TURKISH:    return L"tr";
        case LANG_VIETNAMESE: return L"vi";
        case LANG_INDONESIAN: return L"id";
        case LANG_THAI:       return L"th";
        case LANG_DUTCH:      return L"nl";
        case LANG_POLISH:     return L"pl";
        case LANG_UKRAINIAN:  return L"uk";
        case LANG_ROMANIAN:   return L"ro";
        case LANG_CZECH:      return L"cs";
        case LANG_SWEDISH:    return L"sv";
        case LANG_HUNGARIAN:  return L"hu";
        case LANG_GREEK:      return L"el";
        case LANG_HEBREW:     return L"he";
        case LANG_PERSIAN:    return L"fa";
        case LANG_DANISH:     return L"da";
        case LANG_FINNISH:    return L"fi";
        case LANG_NORWEGIAN:  return L"nb";
        default:              return L"en";
    }
}

static int g_langIndex = 0;

static int IndexForTag(std::wstring_view tag) {
    if (tag.empty()) tag = DetectSystemLangTag();
    auto low = util::Lower(tag);
    if (low == L"zh") low = L"zh-hans";
    if (low == L"no") low = L"nb";
    if (low == L"pt-br" || low == L"pt-pt") low = L"pt";
    for (int i = 0; i < LangCount(); ++i)
        if (util::IEquals(kLangs[i].tag, low)) return i;
    // prefix match: "de-DE" -> "de"
    auto dash = low.find(L'-');
    if (dash != std::wstring_view::npos) {
        auto prim = low.substr(0, dash);
        for (int i = 0; i < LangCount(); ++i)
            if (util::IEquals(kLangs[i].tag, prim)) return i;
    }
    return 0; // English fallback
}

int CurrentLang() { return g_langIndex; }

void SetLanguage(const wchar_t* tag) {
    g_langIndex = IndexForTag(tag ? tag : L"");
}

bool CurrentLangIsRtl() { return kLangs[g_langIndex].rtl; }
bool RtlActiveViaI18n() { return kLangs[g_langIndex].rtl; }

namespace psqc {
const wchar_t* Str(int id) {
    if (id < 0 || id >= STR_COUNT) return L"?";
    const wchar_t* s = kTables[g_langIndex].table[id];
    if (!s || !*s) s = kStr_en[id];
    if (!s) s = L"?";
    return s;
}
} // namespace psqc
