// Platform-independent logic implementation (no Windows APIs).
#include "pure.h"
#include <cstdio>
#include <cstdlib>
#include <cwchar>
#include <vector>

namespace util {

static std::string Utf8FromWchar(wchar_t c) {
    std::string out;
    unsigned cp = (unsigned)c;
    if (cp < 0x80) out += (char)cp;
    else if (cp < 0x800) {
        out += (char)(0xC0 | (cp >> 6));
        out += (char)(0x80 | (cp & 0x3F));
    } else {
        out += (char)(0xE0 | (cp >> 12));
        out += (char)(0x80 | ((cp >> 6) & 0x3F));
        out += (char)(0x80 | (cp & 0x3F));
    }
    return out;
}

std::string JsonEscape(const std::wstring& s) {
    std::string out;
    for (wchar_t c : s) {
        switch (c) {
            case L'"': out += "\\\""; break;
            case L'\\': out += "\\\\"; break;
            case L'\n': out += "\\n"; break;
            case L'\r': out += "\\r"; break;
            case L'\t': out += "\\t"; break;
            default:
                if (c < 0x20) { char b[8]; snprintf(b, sizeof b, "\\u%04x", (unsigned)c); out += b; }
                else out += Utf8FromWchar(c);
        }
    }
    return out;
}

namespace {
struct JsonReader {
    std::string_view s;
    size_t i = 0;
    bool ok = true;
    void SkipWs() { while (i < s.size() && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) ++i; }
    bool Ch(char c) { SkipWs(); if (i < s.size() && s[i] == c) { ++i; return true; } return false; }
    bool Peek(char c) { SkipWs(); return i < s.size() && s[i] == c; }
    std::wstring ParseString() {
        std::wstring out;
        if (!Ch('"')) { ok = false; return out; }
        while (i < s.size() && s[i] != '"') {
            char c = s[i++];
            if (c == '\\' && i < s.size()) {
                char e = s[i++];
                switch (e) {
                    case 'n': out += L'\n'; break;
                    case 't': out += L'\t'; break;
                    case 'r': out += L'\r'; break;
                    case 'b': out += L'\b'; break;
                    case 'f': out += L'\f'; break;
                    case '"': out += L'"'; break;
                    case '\\': out += L'\\'; break;
                    case '/': out += L'/'; break;
                    case 'u': {
                        if (i + 4 <= s.size()) {
                            wchar_t w = 0;
                            for (int k = 0; k < 4; ++k) {
                                char h = s[i++];
                                w = (wchar_t)(w << 4);
                                if (h >= '0' && h <= '9') w = (wchar_t)(w | (h - '0'));
                                else if (h >= 'a' && h <= 'f') w = (wchar_t)(w | (h - 'a' + 10));
                                else if (h >= 'A' && h <= 'F') w = (wchar_t)(w | (h - 'A' + 10));
                            }
                            out += w;
                        } else ok = false;
                        break;
                    }
                    default: ok = false;
                }
            } else {
                unsigned char uc = (unsigned char)c;
                int extra = uc >= 0xF0 ? 3 : uc >= 0xE0 ? 2 : uc >= 0xC0 ? 1 : 0;
                unsigned cp = extra == 0 ? uc : uc & (0x7F >> (extra + 1));
                for (int k = 0; k < extra && i < s.size(); ++k) cp = (cp << 6) | ((unsigned char)s[i++] & 0x3F);
                out += (wchar_t)cp;   // BMP-only; astral chars are not used by our data
            }
        }
        if (i < s.size()) ++i; else ok = false;
        return out;
    }
    void ParseValue(JsonValue& v) {
        SkipWs();
        if (i >= s.size()) { ok = false; return; }
        char c = s[i];
        if (c == '{') {
            ++i;
            while (true) {
                SkipWs();
                if (Peek('}')) { ++i; return; }
                std::wstring key = ParseString();
                if (!ok) return;
                if (!Ch(':')) { ok = false; return; }
                JsonValue child;
                ParseValue(child);
                if (!ok) return;
                v.obj[key] = child;
                if (Ch(',')) continue;
                if (Ch('}')) return;
                ok = false; return;
            }
        } else if (c == '[') {
            ++i;
            while (true) {
                SkipWs();
                if (Peek(']')) { ++i; return; }
                JsonValue child;
                ParseValue(child);
                if (!ok) return;
                v.obj[std::to_wstring(v.obj.size())] = child;
                if (Ch(',')) continue;
                if (Ch(']')) return;
                ok = false; return;
            }
        } else if (c == '"') {
            v.type = JsonValue::Str;
            v.str = ParseString();
        } else if (c == 't' && s.substr(i, 4) == "true") { i += 4; v.type = JsonValue::Bool; v.b = true; }
        else if (c == 'f' && s.substr(i, 5) == "false") { i += 5; v.type = JsonValue::Bool; v.b = false; }
        else if (c == 'n' && s.substr(i, 4) == "null") { i += 4; }
        else {
            size_t b = i;
            while (i < s.size() && (isdigit((unsigned char)s[i]) || s[i] == '-' || s[i] == '+' || s[i] == '.' || s[i] == 'e' || s[i] == 'E')) ++i;
            if (b == i) { ok = false; return; }
            v.type = JsonValue::Num;
            v.num = atof(std::string(s.substr(b, i - b)).c_str());
        }
    }
};
} // namespace

bool JsonParse(std::string_view text, JsonValue& out) {
    JsonReader r{ text, 0, true };
    r.ParseValue(out);
    return r.ok;
}

static std::vector<int> ParseVer(std::wstring_view s) {
    std::vector<int> parts;
    // trim
    size_t b = s.find_first_not_of(L" \t\r\n");
    if (b == std::wstring_view::npos) b = s.size();
    size_t e = s.find_last_not_of(L" \t\r\n");
    std::wstring t = (e == std::wstring_view::npos) ? std::wstring() : std::wstring(s.substr(b, e - b + 1));
    if (!t.empty() && (t[0] == L'v' || t[0] == L'V')) t = t.substr(1);
    size_t start = 0;
    while (parts.size() < 4) {
        size_t p = t.find(L'.', start);
        std::wstring chunk = (p == std::wstring::npos) ? t.substr(start) : t.substr(start, p - start);
        std::wstring digits;
        for (wchar_t c : chunk) { if (c >= L'0' && c <= L'9') digits += c; else break; }
        parts.push_back(digits.empty() ? 0 : (int)wcstol(digits.c_str(), nullptr, 10));
        if (p == std::wstring::npos) break;
        start = p + 1;
    }
    while (parts.size() < 4) parts.push_back(0);
    return parts;
}

int CompareVersions(std::wstring_view a, std::wstring_view b) {
    auto va = ParseVer(a), vb = ParseVer(b);
    for (size_t i = 0; i < 4; ++i) {
        if (va[i] != vb[i]) return va[i] < vb[i] ? -1 : 1;
    }
    return 0;
}

} // namespace util
