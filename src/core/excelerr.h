#pragma once
#include <cstdint>

// Excel's error values, named once. The same seven arrive as xlerrNull 0 .. xlerrNA 42 from the C
// API and as those plus 2000 from VBA (xlErrNA = 2042). A code outside the set gets no name, and
// the caller keeps the raw number.

namespace core
{
    // The published set. Anything else returns nullptr: say the number instead.
    // One table: the spelling a cell shows, and the short form the caller column uses.
    struct ExcelErrEntry { int code; const char* name; const char* shortName; };

    inline const ExcelErrEntry* FindExcelErr(int e)
    {
        static const ExcelErrEntry kTable[] = {
            { 0,  "#NULL!",        "null"  },
            { 7,  "#DIV/0!",       "div0"  },
            { 15, "#VALUE!",       "value" },
            { 23, "#REF!",         "ref"   },
            { 29, "#NAME?",        "name"  },
            { 36, "#NUM!",         "num"   },
            { 42, "#N/A",          "na"    },
            { 43, "#GETTING_DATA", nullptr },
        };
        for (const ExcelErrEntry& x : kTable) if (x.code == e) return &x;
        return nullptr;
    }

    inline const char* ExcelErrName(int e)
    {
        const ExcelErrEntry* x = FindExcelErr(e);
        return x ? x->name : nullptr;
    }

    inline const char* ExcelErrShortName(int e)
    {
        const ExcelErrEntry* x = FindExcelErr(e);
        return x ? x->shortName : nullptr;
    }

    // VBA's numbering as it sits in a VARIANT's SCODE: CVErr(2042) stores 0x800A07FA,
    // FACILITY_CONTROL with 2042 in the low word. The bare number is accepted too; it cannot
    // collide, since it would be a success code. DISP_E_PARAMNOTFOUND is facility 0x2, so an
    // omitted Optional is never named as a cell error.
    inline const char* ExcelErrNameFromVba(std::uint32_t scode)
    {
        std::uint32_t n = 0;
        if ((scode & 0xFFFF0000u) == 0x800A0000u) n = scode & 0xFFFFu;
        else if (scode <= 0xFFFFu)                n = scode;
        else                                      return nullptr;

        if (n < 2000u || n > 2043u) return nullptr;
        return ExcelErrName(static_cast<int>(n - 2000u));
    }
}
