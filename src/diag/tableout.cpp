#include "tableout.h"

#include "core/excel_api.h"

#include <windows.h>
#include <cwchar>

namespace diag
{
namespace
{
    std::string Utf8(const std::wstring& w)
    {
        if (w.empty()) return std::string();
        const int n = WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()),
                                          nullptr, 0, nullptr, nullptr);
        if (n <= 0) return std::string();
        std::string out(static_cast<size_t>(n), '\0');
        WideCharToMultiByte(CP_UTF8, 0, w.data(), static_cast<int>(w.size()), &out[0], n, nullptr, nullptr);
        return out;
    }

    void AppendCsvField(std::string& to, const std::wstring& value)
    {
        const std::string s = Utf8(value);
        if (s.find_first_of(",\"\r\n") == std::string::npos) { to += s; return; }
        to += '"';
        for (char c : s)
        {
            to += c;
            if (c == '"') to += '"';
        }
        to += '"';
    }

}

std::string AsCsv(const Table& t)
{
    std::string out = "\xEF\xBB\xBF";
    for (size_t c = 0; c < t.columns.size(); ++c)
    {
        if (c) out += ',';
        AppendCsvField(out, t.columns[c].title);
    }
    out += "\r\n";
    for (const std::vector<std::wstring>& row : t.rows)
    {
        for (size_t c = 0; c < t.columns.size(); ++c)
        {
            if (c) out += ',';
            if (c < row.size()) AppendCsvField(out, row[c]);
        }
        out += "\r\n";
    }
    return out;
}

std::wstring AsTabbed(const Table& t)
{
    std::wstring out;
    for (size_t c = 0; c < t.columns.size(); ++c)
    {
        if (c) out += L'\t';
        out += t.columns[c].title;
    }
    out += L"\r\n";
    for (const std::vector<std::wstring>& row : t.rows)
    {
        for (size_t c = 0; c < t.columns.size(); ++c)
        {
            if (c) out += L'\t';
            if (c < row.size()) out += row[c];
        }
        out += L"\r\n";
    }
    return out;
}

bool WriteBytes(const std::wstring& path, const std::string& bytes)
{
    HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr,
                           CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
    if (h == INVALID_HANDLE_VALUE) return false;
    DWORD written = 0;
    const BOOL ok = ::WriteFile(h, bytes.data(), static_cast<DWORD>(bytes.size()), &written, nullptr);
    CloseHandle(h);
    return ok && written == bytes.size();
}

std::wstring ExportPath(const wchar_t* pageName)
{
    const std::wstring folder = core::EnsureAppSubdir(L"Diagnostics");
    if (folder.empty()) return L"";
    SYSTEMTIME now{};
    GetLocalTime(&now);
    wchar_t name[MAX_PATH];
    _snwprintf_s(name, _TRUNCATE, L"%s\\XRayXL_%s_%lu_%04u%02u%02u-%02u%02u%02u.csv",
                 folder.c_str(), pageName ? pageName : L"Diagnostics", GetCurrentProcessId(),
                 now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, now.wSecond);
    return name;
}
}
