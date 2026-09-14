// WHO CALLED THIS? -- decoding xlfCaller, all of it.
//
// A cell is only one of the things xlfCaller can name, and the others are not
// errors but different, true answers:
//
//   a cell formula          -> a REFERENCE (xltypeSRef, or xltypeRef for a
//                              multi-cell / array-formula caller)
//   a macro on a button,
//   a shape, a picture      -> a STRING: the graphic object's name
//   a toolbar tool          -> an ARRAY of two numbers, {toolbar, position}
//   a menu command          -> an ARRAY of four, {bar ID, menu, submenu, command}
//   the macro dialog, an
//   Auto macro, an event,
//   the VBE, DDE/OLE        -> the ERROR #REF! -- there is no caller on a sheet
//   Excel calling for its
//   own purposes            -> a reference at row 0, column 0, with NO sheet
//
// THAT LAST ONE IS A TRAP: row 0 column 0 decodes arithmetically to a
// confident "A1". A real calculation names a real sheet, so a cell is written
// ONLY when the sheet resolves too.
//
// ONE DECODER FOR BOTH SIDES, because two would be two chances to decode that
// trap differently -- and a producer and a consumer that do not share a decoder
// verify nothing about what ships.
//
// EVERY OUTCOME IS NAMED, INCLUDING THE REFUSALS: `kind` is never empty --
// "cell", "object:Button 1", "none:ref", "unavailable:2". An instrument that
// cannot tell "never asked" from "asked and Excel declined" reports the second
// as the first.
#pragma once

namespace core
{
    // NEVER ZERO-INITIALISE ONE. `Caller who;` -- not `Caller who{}`. Every
    // path that fills one starts by setting the three first bytes and isCell,
    // and every field is NUL-terminated, so `{}` memsets the whole struct on
    // EVERY TRACED CALL to then write four bytes. On a hot path that is a cost
    // paid for nothing.
    //
    // THAT IS ALSO WHY THE BUFFERS CAN BE GENEROUS. This lives on the stack for
    // the length of one call; an untouched stack byte costs a larger `sub rsp`
    // in the prologue and nothing else. Sizing for a real worst case is free
    // ONCE THE MEMSET IS GONE -- the two decisions look independent and are not.
    // The ceiling is 4 KB: past that MSVC inserts _chkstk probes, which would
    // be a genuine per-call cost.
    struct Caller
    {
        // THE KIND, and the description whose meaning the kind decides. Two
        // fields, because they answer two questions and a reader filtering on
        // one should not have to spell out the other.
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
        // `name`, NOT `object`. A graphic object and an Auto_Open/Auto_Close
        // macro both come back as xltypeStr and Excel offers no way to tell
        // them apart -- the Auto case is the calling SHEET's name. Calling the
        // kind `object` asserts "a shape", which is simply wrong half the time;
        // `name` says what is certain -- Excel returned a name -- and leaves
        // the description to say which name.
        //
        // THREE DIFFERENT "NO" ANSWERS, kept apart deliberately: `none` is
        // Excel saying there is no caller on a sheet, `unavailable` is Excel
        // declining to answer, `not-asked` is us never having asked. An
        // instrument that cannot tell those apart reports the third as the
        // first.
        char kind[24];

        // Sized in BYTES, not characters: a sheet name is capped at 31
        // characters but a workbook name is not, and UTF-8 costs up to four
        // bytes each. Generous because it is free -- see the note above about
        // never zero-initialising one.
        char desc[1024];

        // kind == "cell". Kept as a flag so callers need no strcmp.
        bool isCell;
    };

    // DECODE ONE ANSWER. Pure: no Excel, no globals, no allocation.
    // `sheetName` is the resolved sheet for a reference caller, or null/empty
    // when it did not resolve -- the signal that this is not a real cell.
    //
    // Split out so it can be tested WITHOUT Excel: some answers cannot be
    // produced by automation at all (a graphic-object caller needs a real mouse
    // click), so an offline harness feeds a synthetic XLOPER12 of every kind
    // through it.
    void DecodeCaller(const void* callerOper, const char* sheetName, Caller& out);

    // Ask Excel, then decode. Safe when the answer is unavailable: it says so
    // in `kind` rather than inventing a cell.
    //
    // Only xlfCaller and xlSheetNm, and the reference is never coerced --
    // xlCoerce is the documented hazard here, failing with xlretUncalced on an
    // uncalculated cell. xlfCaller is also the one XLM information function the
    // C API permits from a thread-safe function under multithreaded calc.
    void ReadCaller(Caller& out);
}
