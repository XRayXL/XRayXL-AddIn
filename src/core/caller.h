// Decoding xlfCaller. A cell is only one of the things it can name, and the others are true
// answers, not errors:
//
//   a cell formula          -> a reference (xltypeSRef, or xltypeRef for a
//                              multi-cell / array-formula caller)
//   a macro on a button,
//   a shape, a picture      -> a string: the graphic object's name
//   a toolbar tool          -> an array of two numbers, {toolbar, position}
//   a menu command          -> an array of four, {bar ID, menu, submenu, command}
//   the macro dialog, an
//   Auto macro, an event,
//   the VBE, DDE/OLE        -> the error #REF! -- there is no caller on a sheet
//   Excel calling for its
//   own purposes            -> a reference at row 0, column 0, with no sheet
//
// The last one is a trap: row 0 column 0 decodes to a confident "A1", so a cell is written only
// when the sheet resolves too.
//
// `kind` is never empty, so "never asked" and "asked and Excel declined" stay apart.
#pragma once

namespace core
{
    // `Caller who;`, not `{}`: every path NUL-terminates each field, so `{}` would memset it on every
    // traced call for nothing. Keep it under 4 KB, past which MSVC inserts _chkstk probes.
    struct Caller
    {
        // The kind, and the description whose meaning the kind decides:
        //
        //   cell        "[Book1]Sheet1!B2"  or a whole CSE range "…!B2:D4"
        //   name        "GoButton"          -- see below
        //   toolbar     "5/2" or "\"MyBar\"/2"
        //   menu        "27/27/14/0"        -- four fields, per the SDK
        //   registerid  "42"
        //   none        "ref" | "nil" | "emptyref" | "sheetless-B2" | "nametoolong"
        //   unavailable "2"                 -- Excel declined; the xlret code
        //   not-asked   ""                  -- CALLER=FALSE; we never asked
        //   array       "7"                 -- an array of a size nobody documents
        //   unknown     "0x4"               -- an XLOPER type we have not seen
        //
        // `name`, not `object`: a graphic object and an Auto_Open/Auto_Close macro both come back
        // as xltypeStr, and Excel offers no way to tell them apart.
        //
        // `none` is Excel saying there is no caller on a sheet, `unavailable` is Excel declining to
        // answer, `not-asked` is us never having asked.
        char kind[24];

        // Bytes, not characters: a workbook name has no length cap and UTF-8 takes up to four
        // bytes each. Generous because, never zero-initialised, it costs nothing.
        char desc[1024];

        // kind == "cell". Kept as a flag so callers need no strcmp.
        bool isCell;
    };

    // Decodes one answer. Pure, so it can be tested without Excel: some answers, such as a
    // graphic-object caller, cannot be produced by automation. `sheetName` is null or empty when
    // the sheet did not resolve, which marks a caller that is not a real cell.
    void DecodeCaller(const void* callerOper, const char* sheetName, Caller& out);

    // Asks Excel, then decodes. Only xlfCaller and xlSheetNm, and the reference is never coerced:
    // xlCoerce fails with xlretUncalced on an uncalculated cell.
    void ReadCaller(Caller& out);
}
