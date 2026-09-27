#include "perfettopage.h"

#include <windows.h>
#include <shlwapi.h>
#include <cctype>
#include <cstring>
#include <vector>

#pragma comment(lib, "shlwapi.lib")

namespace ui
{
namespace perfetto
{
namespace
{
    struct Handle
    {
        HANDLE h;
        explicit Handle(HANDLE v) : h(v) {}
        ~Handle() { if (h != INVALID_HANDLE_VALUE) CloseHandle(h); }
        Handle(const Handle&) = delete;
        Handle& operator=(const Handle&) = delete;
    };

    bool Put(HANDLE out, const char* data, size_t n)
    {
        DWORD wrote = 0;
        return WriteFile(out, data, static_cast<DWORD>(n), &wrote, nullptr) && wrote == n;
    }
    bool Put(HANDLE out, const std::string& s) { return Put(out, s.data(), s.size()); }

    // Base64, padded when `n` is not a whole number of three-byte groups.
    void Encode(const unsigned char* in, size_t n, std::string& out)
    {
        static const char kAlphabet[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        out.clear();
        size_t i = 0;
        for (; i + 3 <= n; i += 3)
        {
            const unsigned v = (in[i] << 16) | (in[i + 1] << 8) | in[i + 2];
            out += kAlphabet[v >> 18];
            out += kAlphabet[(v >> 12) & 63];
            out += kAlphabet[(v >> 6) & 63];
            out += kAlphabet[v & 63];
        }
        if (i < n)
        {
            const bool two = i + 1 < n;
            const unsigned v = (in[i] << 16) | (two ? in[i + 1] << 8 : 0);
            out += kAlphabet[v >> 18];
            out += kAlphabet[(v >> 12) & 63];
            out += two ? kAlphabet[(v >> 6) & 63] : '=';
            out += '=';
        }
    }

    // The trace's file name as a JavaScript string: plain ASCII, everything else escaped.
    std::string NameLiteral(const std::wstring& trace)
    {
        const size_t slash = trace.find_last_of(L"\\/");
        const std::wstring name = slash == std::wstring::npos ? trace : trace.substr(slash + 1);
        std::string out = "\"";
        for (wchar_t c : name)
        {
            if (c >= 0x20 && c < 0x7F && c != L'"' && c != L'\\') { out += static_cast<char>(c); continue; }
            char esc[8];
            _snprintf_s(esc, _TRUNCATE, "\\u%04X", static_cast<unsigned>(c));
            out += esc;
        }
        return out + "\"";
    }

    Written Body(const std::wstring& trace, HANDLE in, HANDLE out)
    {
        if (!Put(out, "XRayXLTrace(" + NameLiteral(trace) + ", \"")) return Written::WriteFailed;

        // Encoded in whole groups; a short read carries its last one or two bytes forward.
        std::vector<unsigned char> buf(3 * 65536);
        size_t have = 0;
        std::string text;
        for (;;)
        {
            DWORD got = 0;
            if (!ReadFile(in, buf.data() + have, static_cast<DWORD>(buf.size() - have), &got, nullptr))
                return Written::ReadFailed;
            have += got;
            const size_t whole = got ? have - have % 3 : have;
            Encode(buf.data(), whole, text);
            if (!Put(out, text)) return Written::WriteFailed;
            memmove(buf.data(), buf.data() + whole, have - whole);
            have -= whole;
            if (!got) break;
        }
        return Put(out, "\");\n") ? Written::Ok : Written::WriteFailed;
    }

    std::wstring FileUrl(const std::wstring& path)
    {
        wchar_t url[2048];
        DWORD chars = 2048;
        return SUCCEEDED(UrlCreateFromPathW(path.c_str(), url, &chars, 0)) ? std::wstring(url) : std::wstring();
    }

    // Everything but letters, digits and -_.~ as %XX of its UTF-8, so it survives after the #.
    std::wstring Escaped(const std::wstring& text)
    {
        std::string utf8(text.size() * 3 + 1, '\0');
        const int n = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), static_cast<int>(text.size()),
                                          &utf8[0], static_cast<int>(utf8.size()), nullptr, nullptr);
        utf8.resize(n > 0 ? n : 0);
        std::wstring out;
        for (unsigned char c : utf8)
        {
            if (isalnum(c) || c == '-' || c == '_' || c == '.' || c == '~') { out += static_cast<wchar_t>(c); continue; }
            wchar_t esc[4];
            swprintf_s(esc, L"%%%02X", c);
            out += esc;
        }
        return out;
    }
}

std::wstring ScriptPathFor(const std::wstring& trace)
{
    const size_t slash = trace.find_last_of(L"\\/");
    const size_t dot = trace.find_last_of(L'.');
    const bool ext = dot != std::wstring::npos && (slash == std::wstring::npos || dot > slash);
    return trace.substr(0, ext ? dot : trace.size()) + L".trace.js";
}

Written WriteScript(const std::wstring& trace, const std::wstring& out)
{
    // shared, as a Tail window may be reading the same file
    Handle in(CreateFileW(trace.c_str(), GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                          nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr));
    LARGE_INTEGER size{};
    if (in.h == INVALID_HANDLE_VALUE || !GetFileSizeEx(in.h, &size)) return Written::ReadFailed;
    if (static_cast<unsigned long long>(size.QuadPart) > kMaxTraceBytes) return Written::TooLarge;

    Written r;
    {
        Handle dst(CreateFileW(out.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr));
        if (dst.h == INVALID_HANDLE_VALUE) return Written::WriteFailed;
        r = Body(trace, in.h, dst.h);
    }
    if (r != Written::Ok) DeleteFileW(out.c_str());     // nothing half-written for the page to load
    return r;
}

std::wstring PageAddress(const std::wstring& page, const std::wstring& script)
{
    const std::wstring pageUrl = FileUrl(page), scriptUrl = FileUrl(script);
    if (pageUrl.empty() || scriptUrl.empty()) return std::wstring();
    return pageUrl + L"#" + Escaped(scriptUrl);
}
}
}
