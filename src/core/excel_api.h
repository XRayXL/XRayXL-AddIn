// Shared Excel C API glue: the Excel12 callback plumbing, the Pascal-string
// helper its arguments need, and the per-user output paths.
#pragma once
#include <windows.h>
#include <vector>
#include <string>
#include "xlcall.h"

// Pascal-style XLOPER12 string (str[0] = length, chars follow, no null
// terminator), the format Excel's C API requires. `oper.val.str` points into
// `buf`, so the PascalStr must outlive every use of the oper.
namespace core
{
struct PascalStr
{
    std::vector<XCHAR> buf;
    XLOPER12 oper{};

    // A copy's oper.val.str would still point into the original buf; moving is safe, since the
    // vector hands the buffer over.
    PascalStr() = default;
    PascalStr(const PascalStr&) = delete;
    PascalStr& operator=(const PascalStr&) = delete;
    PascalStr(PascalStr&&) = default;
    PascalStr& operator=(PascalStr&&) = default;
};

PascalStr MakeStr(const wchar_t* text);

// Full path of this DLL, from the address of a function in it, so no DllMain
// HMODULE bookkeeping is needed.
std::wstring GetOwnModulePath();

// %TEMP%\XRayXL\<leaf>, creating both levels. Falls back to the module
// directory if %TEMP% is unusable, so output is never lost.
std::wstring EnsureAppSubdir(const wchar_t* leaf);
// Why XRAYXL_OUTPUT_DIR was not used, or nullptr when it was or was not set.
const char* OutputDirRefused();
}   // namespace core

extern "C"
{
    int __cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count, ...);
}
