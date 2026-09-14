#pragma once
#include "xlcall.h"
#include "core/tracemodes.h"

namespace app
{

// PARSING THE TRACE-PARAM COMMAND SURFACE -- the XLOPER12 -> value
// conversions behind XRayXL_SetTraceParam / GetTraceParam, kept apart from
// commands.cpp so the grammar can be unit-tested without Excel. Nothing here
// calls Excel: these read an
// argument the host already delivered. The state changes and the echo stay in
// commands.cpp.
namespace params
{
    // The argument's type with the free-flags masked off; xltypeMissing for null.
    int  ArgType(LPXLOPER12 v);
    // Missing or nil -- an omitted argument (a leading arg may be omitted).
    bool IsMissing(LPXLOPER12 v);
    // The pascal-counted payload of a string XLOPER12, uppercased into `buf`.
    void ReadUpper(const XLOPER12* v, wchar_t* buf, int cap);
    // Is this argument the (case-insensitive) word `w`?
    bool IsWord(LPXLOPER12 v, const wchar_t* w);

    bool ParseSource  (LPXLOPER12 v, bool& both, core::modes::Source& s);  // XLL|VBA, or omitted=both
    bool ParseParam   (LPXLOPER12 v, core::modes::Param& p);               // DEPTH|ARGS|RETVAL|OBJECTS
    bool ParseDepth   (LPXLOPER12 v, core::modes::Depth& d);               // OFF|TOP|ALL
    bool ParseOnOff   (LPXLOPER12 v, bool& on);                      // TRUE|FALSE (native bool or text)
    bool ParseWhenFull(LPXLOPER12 v, bool& pause);                   // DROP|PAUSE

    // BUFFERSIZE in bytes: bare or M/MB = megabytes, K/KB = kilobytes,
    // case-insensitive; a native number is megabytes. False on a bad number or
    // unit, or above the 4 GB cap. The caller enforces the ring floor below.
    bool ParseBufferBytes(LPXLOPER12 v, unsigned long long& outBytes);

    // The smallest ring accepted (bytes); 0 (synchronous) is the only smaller value.
    // A row too big for the ring is dropped and counted.
    inline constexpr unsigned long long kBufMinRing = 16ull * 1024ull;

    // Render a byte budget the way ParseBufferBytes accepts it back (symmetric):
    // "0", "<n>MB" when it divides evenly, else "<n>KB".
    void FormatBufferW(unsigned long long bytes, wchar_t* out, int cap);
}
}   // namespace app
