#pragma once

#include <string>

// What the Options dialog does with a path. Each can fail, and says how.

namespace ui
{
namespace trace
{
    enum class Result { Ok = 0, NoTraceFile, ClipboardFailed, LaunchFailed };

    Result CopyText(void* ownerHwnd, const std::wstring& text);
    Result OpenFolder(const std::wstring& folder);
    Result RevealFile(const std::wstring& file);

    // A PowerShell window following `file`, so lines appear as they are written.
    Result TailInPowerShell(const std::wstring& file);

    // What to tell a person, for any result. Never null.
    const wchar_t* Explain(Result r);
}
}
