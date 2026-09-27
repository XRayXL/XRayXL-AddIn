// A trace handed to the Perfetto page: the script beside it carries its exact bytes, calls what the
// shipping page defines, and the page's address carries the script's. No Excel and no browser.
#include "ui/perfettopage.h"

#include <windows.h>
#include <cstdio>
#include <string>
#include <vector>

namespace P = ui::perfetto;

static int g_fail = 0;
static int g_pass = 0;

static void Check(const std::string& name, bool ok, const char* detail = "")
{
    if (ok) { ++g_pass; std::printf("  ok   %s\n", name.c_str()); }
    else    { ++g_fail; std::printf("  FAIL %s  %s\n", name.c_str(), detail); }
}

static bool ReadAll(const std::wstring& path, std::string& out)
{
    FILE* f = nullptr;
    if (_wfopen_s(&f, path.c_str(), L"rb") != 0 || !f) return false;
    out.clear();
    char buf[65536];
    size_t n;
    while ((n = fread(buf, 1, sizeof buf, f)) > 0) out.append(buf, n);
    fclose(f);
    return true;
}

static bool WriteAll(const std::wstring& path, const std::string& data)
{
    FILE* f = nullptr;
    if (_wfopen_s(&f, path.c_str(), L"wb") != 0 || !f) return false;
    const bool ok = fwrite(data.data(), 1, data.size(), f) == data.size();
    fclose(f);
    return ok;
}

static bool Exists(const std::wstring& path) { return GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES; }

// A decoder written separately from the encoder under test.
static bool Base64Decode(const std::string& in, std::string& out)
{
    auto val = [](char c) -> int {
        if (c >= 'A' && c <= 'Z') return c - 'A';
        if (c >= 'a' && c <= 'z') return c - 'a' + 26;
        if (c >= '0' && c <= '9') return c - '0' + 52;
        if (c == '+') return 62;
        if (c == '/') return 63;
        return -1;
    };
    if (in.size() % 4) return false;
    out.clear();
    for (size_t i = 0; i < in.size(); i += 4)
    {
        int v[4];
        int pad = 0;
        for (int k = 0; k < 4; ++k)
        {
            if (in[i + k] == '=' && i + 4 == in.size() && k >= 2) { v[k] = 0; ++pad; continue; }
            if (pad || (v[k] = val(in[i + k])) < 0) return false;
        }
        const unsigned n = (v[0] << 18) | (v[1] << 12) | (v[2] << 6) | v[3];
        out += static_cast<char>(n >> 16);
        if (pad < 2) out += static_cast<char>((n >> 8) & 0xFF);
        if (pad < 1) out += static_cast<char>(n & 0xFF);
    }
    return true;
}

static std::wstring Here()
{
    wchar_t exe[MAX_PATH];
    GetModuleFileNameW(nullptr, exe, MAX_PATH);
    std::wstring s = exe;
    return s.substr(0, s.find_last_of(L'\\'));
}

int main()
{
    std::printf("perfettopage_test\n");

    // build\x64\Release\unit\ -> the repository: the page must define what the script calls
    std::string page;
    const bool read = ReadAll(Here() + L"\\..\\..\\..\\..\\perfetto\\XRayXL-Perfetto.html", page);
    Check("shipping-page-defines-XRayXLTrace", read && page.find("window.XRayXLTrace = function") != std::string::npos,
          read ? "the page no longer defines window.XRayXLTrace" : "perfetto\\XRayXL-Perfetto.html not found from the exe");
    Check("shipping-page-loads-a-trace-js", read && page.find("\\.trace\\.js$/i.test(handed)") != std::string::npos,
          "the page no longer loads a .trace.js named after its #");

    wchar_t tmp[MAX_PATH];
    GetTempPathW(MAX_PATH, tmp);
    wchar_t dirName[64];
    swprintf_s(dirName, L"perfettopage_test_%lu", GetCurrentProcessId());
    const std::wstring dir = std::wstring(tmp) + dirName;
    CreateDirectoryW(dir.c_str(), nullptr);
    const std::wstring trace = dir + L"\\XRayXL_Trace_1_2.csv";
    const std::wstring script = P::ScriptPathFor(trace);
    const std::string head = "XRayXLTrace(\"XRayXL_Trace_1_2.csv\", \"", tail = "\");\n";

    // ---- every byte comes back, at every length a three-byte group can end on ----
    for (size_t size : std::vector<size_t>{ 0, 1, 2, 3, 4, 5, 6, 7, 196607, 196608, 196609, 196610, 1000003 })
    {
        std::string data(size, '\0');
        unsigned x = 2463534242u + static_cast<unsigned>(size);
        for (size_t i = 0; i < size; ++i) { x ^= x << 13; x ^= x >> 17; x ^= x << 5; data[i] = static_cast<char>(x); }
        const std::string name = "round-trip-" + std::to_string(size) + "-bytes";
        if (!WriteAll(trace, data)) { Check(name, false, "could not write the trace"); continue; }
        std::string js, decoded;
        const bool written = P::WriteScript(trace, script) == P::Written::Ok && ReadAll(script, js);
        const bool framed = written && js.size() >= head.size() + tail.size() &&
                            js.compare(0, head.size(), head) == 0 &&
                            js.compare(js.size() - tail.size(), tail.size(), tail) == 0;
        const bool same = framed &&
                          Base64Decode(js.substr(head.size(), js.size() - head.size() - tail.size()), decoded) &&
                          decoded == data;
        Check(name, same, !written ? "not written" : !framed ? "not a call to XRayXLTrace" : "the bytes differ");
    }

    // ---- the name is a JavaScript string, plain ASCII ----
    {
        const std::wstring accented = dir + L"\\caf\u00E9.csv";
        WriteAll(accented, "x");
        const std::wstring out = P::ScriptPathFor(accented);
        std::string js;
        const bool ok = P::WriteScript(accented, out) == P::Written::Ok && ReadAll(out, js) &&
                        js.rfind("XRayXLTrace(\"caf\\u00E9.csv\", \"", 0) == 0;
        Check("name-is-escaped-for-javascript", ok, js.substr(0, 40).c_str());
        DeleteFileW(accented.c_str());
        DeleteFileW(out.c_str());
    }

    // ---- refusals write nothing ----
    DeleteFileW(script.c_str());
    Check("missing-trace-is-refused",
          P::WriteScript(dir + L"\\none.csv", script) == P::Written::ReadFailed && !Exists(script));
    {
        const std::wstring big = dir + L"\\XRayXL_Trace_3_4.csv";
        HANDLE h = CreateFileW(big.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        DWORD ignored = 0;
        DeviceIoControl(h, FSCTL_SET_SPARSE, nullptr, 0, nullptr, 0, &ignored, nullptr);
        LARGE_INTEGER end;
        end.QuadPart = static_cast<LONGLONG>(P::kMaxTraceBytes) + 1;
        SetFilePointerEx(h, end, nullptr, FILE_BEGIN);
        SetEndOfFile(h);
        CloseHandle(h);
        const std::wstring bigOut = P::ScriptPathFor(big);
        Check("trace-over-the-limit-is-refused",
              P::WriteScript(big, bigOut) == P::Written::TooLarge && !Exists(bigOut));
        DeleteFileW(big.c_str());
    }

    // ---- where the script goes, and the address that names it ----
    Check("script-path-replaces-csv",
          P::ScriptPathFor(L"C:\\t\\XRayXL_Trace_1_2.csv") == L"C:\\t\\XRayXL_Trace_1_2.trace.js");
    Check("script-path-replaces-jsonl",
          P::ScriptPathFor(L"C:\\t\\XRayXL_Trace_1_2.jsonl") == L"C:\\t\\XRayXL_Trace_1_2.trace.js");
    Check("script-path-ignores-a-dot-in-the-folder",
          P::ScriptPathFor(L"C:\\a.b\\trace") == L"C:\\a.b\\trace.trace.js");
    const std::wstring address = P::PageAddress(L"C:\\x y\\XRayXL-Perfetto.html", L"C:\\t u\\XRayXL_Trace_1_2.trace.js");
    const std::wstring expected =
        L"file:///C:/x%20y/XRayXL-Perfetto.html#file%3A%2F%2F%2FC%3A%2Ft%2520u%2FXRayXL_Trace_1_2.trace.js";
    std::string got;
    for (wchar_t c : address) got += c < 128 ? static_cast<char>(c) : '?';
    Check("page-address-carries-the-script", address == expected, got.c_str());

    DeleteFileW(trace.c_str());
    DeleteFileW(script.c_str());
    RemoveDirectoryW(dir.c_str());

    std::printf("\n%d passed, %d failed\n", g_pass, g_fail);
    return g_fail ? 1 : 0;
}
