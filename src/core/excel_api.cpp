#include "excel_api.h"
#include <cstdarg>

// Excel12 resolves Excel's own MdCallBack12 export -- the standard XLL
// mechanism (Microsoft's XLCALL.CPP sample).
namespace
{
    typedef int(__stdcall* Excel12Proc)(int xlfn, int coper, LPXLOPER12* rgpxloper12, LPXLOPER12 xloper12Res);
    Excel12Proc g_pExcel12 = nullptr;

    void FetchEntryPoint()
    {
        if (g_pExcel12 != nullptr)
            return;
        HMODULE hModule = GetModuleHandle(nullptr);
        if (hModule != nullptr)
            g_pExcel12 = reinterpret_cast<Excel12Proc>(GetProcAddress(hModule, "MdCallBack12"));
    }
}

extern "C" int __cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count, ...)
{
    constexpr int kMax = 32;
    LPXLOPER12 args[kMax];
    FetchEntryPoint();
    if (g_pExcel12 == nullptr) return xlretFailed;
    if (count < 0 || count > kMax) return xlretInvCount;

    va_list ap;
    va_start(ap, count);
    for (int i = 0; i < count; i++) args[i] = va_arg(ap, LPXLOPER12);
    va_end(ap);

    return g_pExcel12(xlfn, count, args, operRes);
}

namespace core
{
PascalStr MakeStr(const wchar_t* text)
{
    PascalStr ps;
    size_t len = wcslen(text);
    ps.buf.resize(len + 1);
    ps.buf[0] = static_cast<XCHAR>(len);
    for (size_t i = 0; i < len; i++) ps.buf[i + 1] = static_cast<XCHAR>(text[i]);
    ps.oper.xltype = xltypeStr;
    ps.oper.val.str = ps.buf.data();
    return ps;
}

// The fallback for EnsureAppSubdir when %TEMP% is unusable.
static std::wstring GetOwnModuleDir()
{
    const std::wstring full = GetOwnModulePath();
    const size_t slash = full.find_last_of(L"\\/");
    return (full.empty() || slash == std::wstring::npos) ? L"." : full.substr(0, slash);
}

namespace
{
    // CreateDirectoryW makes only the last component, and the override below
    // can name a folder several levels deep that does not exist yet.
    bool CreateDirectoryTreeW(const std::wstring& path)
    {
        if (path.empty()) return false;
        if (CreateDirectoryW(path.c_str(), nullptr)) return true;
        const DWORD err = GetLastError();
        if (err == ERROR_ALREADY_EXISTS) return true;
        if (err != ERROR_PATH_NOT_FOUND) return false;

        const size_t slash = path.find_last_of(L"\\/");
        if (slash == std::wstring::npos || slash == 0) return false;
        if (!CreateDirectoryTreeW(path.substr(0, slash))) return false;
        return CreateDirectoryW(path.c_str(), nullptr) ||
               GetLastError() == ERROR_ALREADY_EXISTS;
    }
}

namespace
{
    const char* g_outputDirRefused = nullptr;   // why XRAYXL_OUTPUT_DIR was not used
}
const char* OutputDirRefused() { return g_outputDirRefused; }

// The one place that resolves the output root: a second copy of the rule could send the trace
// file somewhere the suites do not look.
std::wstring EnsureAppSubdir(const wchar_t* leaf)
{
    // An explicit output root, so a suite can keep each run's trace and logs beside its test. Not
    // an override of %TEMP%, which would move Excel's own temporary files too.
    wchar_t over[MAX_PATH];
    const DWORD ov = GetEnvironmentVariableW(L"XRAYXL_OUTPUT_DIR", over, MAX_PATH);
    if (ov >= MAX_PATH) g_outputDirRefused = "it is MAX_PATH characters or longer";
    if (ov > 0 && ov < MAX_PATH)
    {
        std::wstring root(over, ov);
        while (!root.empty() && (root.back() == L'\\' || root.back() == L'/'))
            root.pop_back();
        const std::wstring full = root + L"\\" + leaf;
        if (CreateDirectoryTreeW(full))
            return full;
        // Bad path or no permission: fall through to %TEMP% rather than lose
        // the output entirely, and say so once the log is open.
        g_outputDirRefused = "its folder could not be created";
    }

    wchar_t tmp[MAX_PATH];
    DWORD n = GetTempPathW(MAX_PATH, tmp);   // includes a trailing backslash
    if (n == 0 || n > MAX_PATH)
        return GetOwnModuleDir();            // fall back rather than lose output

    std::wstring xray = std::wstring(tmp, n) + L"XRayXL";
    CreateDirectoryW(xray.c_str(), nullptr);        // ok if it already exists
    std::wstring full = xray + L"\\" + leaf;
    if (!CreateDirectoryW(full.c_str(), nullptr) && GetLastError() != ERROR_ALREADY_EXISTS)
        return GetOwnModuleDir();
    return full;
}

std::wstring GetOwnModulePath()
{
    HMODULE hModule = nullptr;
    GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(&GetOwnModulePath), &hModule);
    wchar_t path[MAX_PATH];
    DWORD len = GetModuleFileNameW(hModule, path, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) return L"";
    return std::wstring(path, len);
}
}   // namespace core
