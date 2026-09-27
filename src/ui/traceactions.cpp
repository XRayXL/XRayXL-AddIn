#include "traceactions.h"
#include "perfettopage.h"

#include "core/log.h"

#include <windows.h>
#include <shellapi.h>
#include <shlwapi.h>
#include <cstring>

#pragma comment(lib, "shlwapi.lib")

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

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

Result OpenInPerfetto(const std::wstring& path)
{
    if (path.empty()) return Result::NoTraceFile;

    // The page ships beside the add-in.
    wchar_t self[MAX_PATH];
    const DWORD n = GetModuleFileNameW(reinterpret_cast<HMODULE>(&__ImageBase), self, MAX_PATH);
    std::wstring page(self, n);
    page = page.substr(0, page.find_last_of(L'\\') + 1) + L"XRayXL-Perfetto.html";
    if (GetFileAttributesW(page.c_str()) == INVALID_FILE_ATTRIBUTES)
    {
        char line[MAX_PATH + 48];
        _snprintf_s(line, _TRUNCATE, "perfetto: %ls is missing", page.c_str());
        core::Log::Warning(line);
        return Result::PageMissing;
    }

    const std::wstring script = perfetto::ScriptPathFor(path);
    const perfetto::Written w = perfetto::WriteScript(path, script);
    if (w == perfetto::Written::TooLarge) return Result::TooLarge;
    if (w != perfetto::Written::Ok)
    {
        char line[96];
        _snprintf_s(line, _TRUNCATE, "perfetto: the trace could not be written for the page (reason %d, error %lu)",
                    static_cast<int>(w), GetLastError());
        core::Log::Warning(line);
        return Result::ScriptFailed;
    }

    // The browser that opens web links: an .html file may open in an editor.
    wchar_t browser[MAX_PATH];
    DWORD chars = MAX_PATH;
    if (FAILED(AssocQueryStringW(ASSOCF_IS_PROTOCOL, ASSOCSTR_EXECUTABLE, L"http", L"open", browser, &chars)))
    {
        core::Log::Warning("perfetto: no default browser is set");
        return Result::LaunchFailed;
    }
    return Launch(browser, L"\"" + perfetto::PageAddress(page, script) + L"\"");
}

const wchar_t* Explain(Result r)
{
    switch (r)
    {
    case Result::Ok:              return L"";
    case Result::NoTraceFile:     return L"The trace file is named when tracing is armed. Arm first, then try again.";
    case Result::ClipboardFailed: return L"The clipboard could not be opened. Another program may be holding it; try again.";
    case Result::TooLarge:        return L"This trace is over 256 MB, too large to hand to the Perfetto page. "
                                         L"Open XRayXL-Perfetto.html, beside the add-in, and drop the trace file on it.";
    case Result::PageMissing:     return L"XRayXL-Perfetto.html is not beside the add-in. It comes with XRayXL: "
                                         L"put it in the same folder as XRayXL64.xll.";
    case Result::ScriptFailed:    return L"The trace could not be written out for the Perfetto page. Details are in the log, "
                                         L"under %TEMP%\\XRayXL\\Logs.";
    default:                      return L"That program could not be started.";
    }
}
}
}
