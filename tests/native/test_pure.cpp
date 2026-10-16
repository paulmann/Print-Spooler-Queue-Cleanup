// Native (POSIX or Windows) unit tests for platform-independent logic:
// i18n table integrity, semver compare, JSON mini-parser.
#include "pure.h"
#include "i18n_ids.h"
#include <cstdio>
#include <cstring>
#include <set>
#include <string>

static int g_fail = 0;
#define CHECK(cond, msg) do { if (!(cond)) { std::printf("FAIL: %s (line %d)\n", msg, __LINE__); ++g_fail; } } while (0)

struct TableRef { const char* tag; const wchar_t* const* t; };
static const TableRef kTables[] = {
    { "en", kStr_en }, { "ru", kStr_ru }, { "es", kStr_es }, { "zh-Hans", kStr_zhHans },
    { "zh-Hant", kStr_zhHant }, { "ar", kStr_ar }, { "hi", kStr_hi }, { "pt", kStr_pt },
    { "fr", kStr_fr }, { "de", kStr_de }, { "ja", kStr_ja }, { "ko", kStr_ko },
    { "it", kStr_it }, { "tr", kStr_tr }, { "vi", kStr_vi }, { "id", kStr_id },
    { "th", kStr_th }, { "nl", kStr_nl }, { "pl", kStr_pl }, { "uk", kStr_uk },
    { "ro", kStr_ro }, { "cs", kStr_cs }, { "sv", kStr_sv }, { "hu", kStr_hu },
    { "el", kStr_el }, { "he", kStr_he }, { "fa", kStr_fa }, { "da", kStr_da },
    { "fi", kStr_fi }, { "nb", kStr_nb },
};

static int CountSub(const wchar_t* hay, const wchar_t* needle) {
    int n = 0;
    const wchar_t* p = hay;
    while ((p = wcsstr(p, needle)) != nullptr) { ++n; ++p; }
    return n;
}

static void TestI18n() {
    CHECK(sizeof(kTables) / sizeof(kTables[0]) == 30, "30 languages registered");
    std::set<std::wstring> natives;
    for (const auto& tr : kTables) {
        for (int id = 0; id < STR_COUNT; ++id) {
            const wchar_t* s = tr.t[id];
            char msg[128];
            snprintf(msg, sizeof msg, "%s: id %d non-empty", tr.tag, id);
            CHECK(s != nullptr && s[0] != 0, msg);
            if (id != STR_LANG_NATIVE && s && kStr_en[id]) {
                int a = CountSub(kStr_en[id], L"%s"), b = CountSub(s, L"%s");
                snprintf(msg, sizeof msg, "%s: id %d %%s-count matches en", tr.tag, id);
                CHECK(a == b, msg);
                int pa = CountSub(kStr_en[id], L"%%"), pb = CountSub(s, L"%%");
                snprintf(msg, sizeof msg, "%s: id %d %%%% -count matches en", tr.tag, id);
                CHECK(pa == pb, msg);
            }
        }
        if (tr.t[STR_LANG_NATIVE]) natives.insert(tr.t[STR_LANG_NATIVE]);
    }
    CHECK(natives.size() == 30, "native names unique across languages");
}

static void TestSemver() {
    CHECK(util::CompareVersions(L"5.0.0", L"5.0.0") == 0, "equal");
    CHECK(util::CompareVersions(L"v5.0.1", L"5.0.0") > 0, "v-prefix greater");
    CHECK(util::CompareVersions(L"5.0.0", L"5.0.1") < 0, "patch less");
    CHECK(util::CompareVersions(L"4.9.9", L"5.0.0") < 0, "major less");
    CHECK(util::CompareVersions(L"5.0.0.1", L"5.0.0") > 0, "quad greater");
    CHECK(util::CompareVersions(L"garbage", L"5.0.0") < 0, "garbage less");
    CHECK(util::CompareVersions(L"5.10.0", L"5.9.9") > 0, "numeric not lexical");
}

static void TestJson() {
    util::JsonValue v;
    bool ok = util::JsonParse("{\"a\": 1, \"b\": \"текст\\n\", \"c\": true, \"d\": [1,2,{\"e\": null}]}", v);
    CHECK(ok, "parse ok");
    CHECK(v.Find(L"a") && v.Find(L"a")->num == 1, "number");
    CHECK(v.Find(L"b") && v.Find(L"b")->str == L"текст\n", "unicode + escape");
    CHECK(v.Find(L"c") && v.Find(L"c")->b, "bool");
    auto* d = v.Find(L"d");
    CHECK(d && d->obj.size() == 3, "array as map");
    CHECK(d && d->Find(L"2") && d->Find(L"2")->Find(L"e") && d->Find(L"2")->Find(L"e")->type == util::JsonValue::Null, "nested null");
    util::JsonValue bad;
    CHECK(!util::JsonParse("{\"a\": ", bad), "truncated fails");
    CHECK(util::JsonEscape(L"a\"b\\c\n") == "a\\\"b\\\\c\\n", "escape");
    util::JsonValue cfg;
    ok = util::JsonParse("{\"themeMode\": 2, \"lang\": \"ru\", \"trayIcon\": false}", cfg);
    CHECK(ok && cfg.Find(L"themeMode")->num == 2 && cfg.Find(L"lang")->str == L"ru" && !cfg.Find(L"trayIcon")->b, "settings-like doc");
}

int main() {
    TestI18n();
    TestSemver();
    TestJson();
    if (g_fail == 0) { std::printf("ALL NATIVE TESTS PASSED (30 languages x %d strings, semver, json)\n", STR_COUNT); return 0; }
    std::printf("%d FAILURES\n", g_fail);
    return 1;
}
