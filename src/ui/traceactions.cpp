#include "traceactions.h"

#include "core/log.h"

#include <windows.h>
#include <shellapi.h>
#include <cstring>

namespace ui
{
namespace trace
{
namespace
{
    Result Launch(const wchar_t* exe, const std::wstring& args)
    {
        // ShellExecute answers 32 or less for failure
        const HINSTANCE rc = ShellExecuteW(nullptr, L"open", exe, args.c_str(), nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(rc) > 32) return Result::Ok;
        char line[96];
        _snprintf_s(line, _TRUNCATE, "options: could not start %ls (code %lld)", exe,
                    static_cast<long long>(reinterpret_cast<INT_PTR>(rc)));
        core::Log::Warning(line);
        return Result::LaunchFailed;
    }
}

Result CopyText(void* ownerHwnd, const std::wstring& text)
{
    if (text.empty()) return Result::NoTraceFile;
    const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
    HGLOBAL mem = GlobalAlloc(GMEM_MOVEABLE, bytes);
    if (!mem) return Result::ClipboardFailed;
    if (void* dst = GlobalLock(mem))
    {
        memcpy(dst, text.c_str(), bytes);
        GlobalUnlock(mem);
    }
    if (OpenClipboard(static_cast<HWND>(ownerHwnd)))
    {
        EmptyClipboard();
        const bool ok = SetClipboardData(CF_UNICODETEXT, mem) != nullptr;   // the clipboard owns it now
        CloseClipboard();
        if (ok) return Result::Ok;
    }
    GlobalFree(mem);
    return Result::ClipboardFailed;
}

Result OpenFolder(const std::wstring& folder)
{
    return folder.empty() ? Result::NoTraceFile : Launch(L"explorer.exe", L"\"" + folder + L"\"");
}

Result RevealFile(const std::wstring& file)
{
    return file.empty() ? Result::NoTraceFile : Launch(L"explorer.exe", L"/select,\"" + file + L"\"");
}

Result TailInPowerShell(const std::wstring& path)
{
    if (path.empty()) return Result::NoTraceFile;

    // single-quoted for PowerShell, with embedded quotes doubled
    std::wstring quoted;
    for (wchar_t ch : path) { quoted.push_back(ch); if (ch == L'\'') quoted.push_back(ch); }

    // -NoExit keeps the window after Ctrl+C. The first row creates the file, so wait for it. The
    // title names the file, and Windows Terminal shows it on the tab.
    const std::wstring args = L"-NoExit -NoProfile -Command \"$p='" + quoted +
        L"'; $Host.UI.RawUI.WindowTitle = 'XRayXL tail: ' + [IO.Path]::GetFileName($p); "
        L"if (-not (Test-Path -LiteralPath $p)) "
        L"{ Write-Host 'Waiting for the trace file...' -ForegroundColor DarkGray; "
        L"while (-not (Test-Path -LiteralPath $p)) { Start-Sleep -Milliseconds 300 } }; "
        L"Get-Content -LiteralPath $p -Tail 40 -Wait\"";
    return Launch(L"powershell.exe", args);
}

const wchar_t* Explain(Result r)
{
    switch (r)
    {
    case Result::Ok:              return L"";
    case Result::NoTraceFile:     return L"The trace file is named when tracing is armed. Arm first, then try again.";
    case Result::ClipboardFailed: return L"The clipboard could not be opened. Another program may be holding it; try again.";
    default:                      return L"That program could not be started.";
    }
}
}
}
