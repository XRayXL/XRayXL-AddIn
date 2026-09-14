#pragma once
#include <cstdint>

// THE SPELLINGS OF EXCEL'S ERROR VALUES, in one place, because the same seven
// errors arrive through two different numberings and were named twice.
//
//   the C API        xlerrNull 0, xlerrDiv0 7 ... xlerrNA 42   (xlcall.h)
//   VBA / the object model    the same, plus 2000              (xlErrNA = 2042)
//
// Two entry points over ONE table of names, so an error reads the same whichever
// side of the tracer saw it: a `#N/A` returned by an XLL function and a `#N/A`
// handed to a VBA UDF are the same fact about the same sheet.
//
// `#ERR?` is NOT in here. A code outside the set is not an Excel error and gets
// no name -- the caller keeps the raw number, which is a true answer where an
// invented name would not be.

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

    // VBA's numbering, as it actually sits in a VARIANT's SCODE.
    //
    // `CVErr(2042)` does NOT store 2042: it stores 0x800A07FA -- FACILITY_CONTROL
    // (0x800A0000) with 2042 in the low word. [measured: a #N/A handed to a
    // Variant parameter, suites/stress/calc/variant-arg-holds-a-cell-error]
    //
    // The bare form is accepted too, for a path that carries the raw xlCVError
    // rather than an SCODE. It cannot collide: an HRESULT of 2042 has its high
    // bit clear, so it is a SUCCESS code, and a VT_ERROR holding a success code
    // is not a thing Excel produces.
    //
    // Everything else keeps its number. DISP_E_PARAMNOTFOUND is 0x80020004 --
    // facility 0x2, not 0xA -- so an omitted Optional can never be named as a
    // cell error, which is the confusion this must not create.
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
