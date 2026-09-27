#pragma once

#include <string>

// A trace handed to the Perfetto page. A page cannot read a file from disk, but it can load a
// script, so the trace is written as one beside it and the page is opened with its address.

namespace ui
{
namespace perfetto
{
    // A browser holds the trace as one base64 string, so a larger one is refused.
    const unsigned long long kMaxTraceBytes = 256ull << 20;

    enum class Written { Ok = 0, TooLarge, ReadFailed, WriteFailed };

    // Beside the trace, named after it: XRayXL_Trace_<id>_<pid>.trace.js.
    std::wstring ScriptPathFor(const std::wstring& trace);

    // Writes `XRayXLTrace("<the trace's file name>", "<its bytes, as base64>");` to `out`.
    Written WriteScript(const std::wstring& trace, const std::wstring& out);

    // The page's file: address, with the script's after the # for the page to load.
    std::wstring PageAddress(const std::wstring& page, const std::wstring& script);
}
}
