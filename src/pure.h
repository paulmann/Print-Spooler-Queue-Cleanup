#pragma once
// Platform-independent logic (compiles on Windows and POSIX for unit tests):
// JSON mini-parser, semantic version compare.
#include <string>
#include <string_view>
#include <map>

namespace util {

struct JsonValue {
    enum Type { Null, Bool, Num, Str } type = Null;
    bool b = false;
    double num = 0;
    std::wstring str;
    std::map<std::wstring, JsonValue> obj;
    const JsonValue* Find(const wchar_t* key) const {
        auto it = obj.find(key);
        return it == obj.end() ? nullptr : &it->second;
    }
};

bool JsonParse(std::string_view text, JsonValue& out);
std::string JsonEscape(const std::wstring& s);
// Returns <0 / 0 / >0 ; parses "v1.2.3", "1.2.3.4", garbage-tolerant.
int CompareVersions(std::wstring_view a, std::wstring_view b);

} // namespace util
