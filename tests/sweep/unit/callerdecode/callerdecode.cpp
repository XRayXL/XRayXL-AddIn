// Everything xlfCaller can say, decoded without Excel.
//
// A live probe proves the common callers: a cell formula, a macro, an event and
// Application.Run. It cannot reach the rest: a graphic-object caller needs a real mouse click
// on a shape, and Application.Run on a button's macro is not a click. So src/core/caller.cpp's
// DecodeCaller is pure, and this feeds it a synthetic XLOPER12 of every kind.
//
// A unit test of a decoder, not evidence about Excel: what each xltype means is Microsoft's
// documentation.
//
// Built by XRayXL.sln into build\x64\Release\unit\.
#include <windows.h>
#include <cstdio>
#include <cstring>
#include <vector>
#include "xlcall.h"
#include "caller.h"

// ReadCaller is not exercised here and cannot be: Excel12 is imported from
// the host process, and there is no host. The stub exists only to let the
// translation unit link, and it returns "failed" so that anything which did
// reach it would be loudly wrong rather than quietly plausible.
extern "C" int __cdecl Excel12(int, LPXLOPER12, int, ...) { return xlretFailed; }
extern "C" int __stdcall Excel12v(int, LPXLOPER12, int, LPXLOPER12[]) { return xlretFailed; }

static int g_fail = 0;

static std::vector<XCHAR> Pascal(const wchar_t* text)
{
    const size_t n = wcslen(text);
    std::vector<XCHAR> v(n + 1);
    v[0] = static_cast<XCHAR>(n);
    for (size_t i = 0; i < n; ++i) v[i + 1] = text[i];
    return v;
}

// Kind and description, the two columns the trace carries, plus isCell.
static void Check(const char* label, const XLOPER12& op, const char* sheet,
                  const char* wantKind, const char* wantDesc)
{
    core::Caller c;
    core::DecodeCaller(&op, sheet, c);

    const bool okKind = std::strcmp(c.kind, wantKind) == 0;
    const bool okDesc = std::strcmp(c.desc, wantDesc) == 0;
    const bool okFlag = c.isCell == (std::strcmp(c.kind, "cell") == 0);

    if (okKind && okDesc && okFlag)
    {
        std::printf("  ok    %-34s %-11s %s\n", label, c.kind, c.desc);
        return;
    }
    ++g_fail;
    std::printf("  FAIL  %-34s kind='%s' (want '%s')  desc='%s' (want '%s')%s\n",
                label, c.kind, wantKind, c.desc, wantDesc,
                okFlag ? "" : "  [isCell disagrees with kind]");
}

int main()
{
    std::printf("decoding every xlfCaller answer, no Excel involved\n\n");

    // ---- a cell formula: the Stage 2 case --------------------------------
    {
        XLOPER12 op{}; op.xltype = xltypeSRef;
        op.val.sref.count = 1;
        op.val.sref.ref.rwFirst = 0;  op.val.sref.ref.rwLast = 0;
        op.val.sref.ref.colFirst = 0; op.val.sref.ref.colLast = 0;
        Check("cell A1", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!A1");

        op.val.sref.ref.rwFirst = 41; op.val.sref.ref.colFirst = 27;
        Check("cell AB42", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!AB42");

        // Bijective base-26: column 26 is Z and 27 is AA. A plain base-26
        // conversion produces "A@" or "AA" one column early.
        op.val.sref.ref.rwFirst = 0; op.val.sref.ref.colFirst = 25;
        Check("cell Z1 (26th column)", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!Z1");
        op.val.sref.ref.colFirst = 26;
        Check("cell AA1 (27th column)", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!AA1");

        // An apostrophe in the sheet name is doubled inside the quotes, as Excel writes it.
        op.val.sref.ref.colFirst = 0;
        Check("sheet name with an apostrophe", op, "[Book1]It's", "cell", "'[Book1]It''s'!A1");
    }

    // The trap: a reference with no sheet. Excel calls functions for its own purposes with no
    // calling cell, and those arrive as row 0 column 0, which must not decode as "A1".
    {
        XLOPER12 op{}; op.xltype = xltypeSRef;
        op.val.sref.count = 1;
        Check("row0/col0, no sheet", op, "", "none", "sheetless-A1");
        Check("row0/col0, null sheet", op, nullptr, "none", "sheetless-A1");
    }

    // ---- a multi-cell / external reference -------------------------------
    {
        XLMREF12 mref{};
        mref.count = 1;
        mref.reftbl[0].rwFirst = 4; mref.reftbl[0].colFirst = 2;
        XLOPER12 op{}; op.xltype = xltypeRef;
        op.val.mref.lpmref = &mref;
        op.val.mref.idSheet = 0;
        // One area, ONE CELL: last == first. Named for what it is -- it was
        // called "array-formula ref" while testing a single cell, which is how
        // the multi-cell case went untested for so long.
        Check("xltypeRef, single cell C5", op, "[Book1]Sheet2", "cell", "[Book1]Sheet2!C5");

        // One area, many cells: what a CSE array formula produces. Both corners must be
        // decoded.
        mref.reftbl[0].rwLast = 6; mref.reftbl[0].colLast = 4;
        Check("xltypeRef, multi-cell C5:E7", op, "[Book1]Sheet2", "cell", "[Book1]Sheet2!C5:E7");
        mref.reftbl[0].rwLast = 0; mref.reftbl[0].colLast = 0;

        // A reference carrying no entries names no cell -- and must still
        // say SOMETHING. This case caught `what` being left empty, which in a
        // trace is indistinguishable from never having asked.
        mref.count = 0;
        Check("ref with count 0", op, "[Book1]Sheet2", "none", "emptyref");
    }

    // ---- A BUTTON. The case a live probe cannot reach. -------------------
    {
        std::vector<XCHAR> name = Pascal(L"GoButton");
        XLOPER12 op{}; op.xltype = xltypeStr;
        op.val.str = name.data();
        Check("graphic object", op, nullptr, "name", "GoButton");

        std::vector<XCHAR> empty = Pascal(L"");
        op.val.str = empty.data();
        Check("graphic object, no name", op, nullptr, "name", "(unnamed)");

        op.val.str = nullptr;
        Check("graphic object, null str", op, nullptr, "name", "(unnamed)");
    }

    // ---- THE REMAINING ROWS OF THE SDK TABLE -----------------------------
    // Each reduces to a shape handled above, but each is a separate ROW in
    // and so gets its own case: the table and
    // this file should map one to one, or a row can quietly have no test.
    {
        // "A conditional formatting expression | A reference to the cell to
        // which the formatting condition is applied."
        XLREF12 sref{}; sref.rwFirst = 1; sref.rwLast = 1;
        sref.colFirst = 1; sref.colLast = 1;
        XLOPER12 op{}; op.xltype = xltypeSRef; op.val.sref.count = 1; op.val.sref.ref = sref;
        Check("conditional formatting expression", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!B2");

        // "A command associated with an xlcOnDoubleclick ... | The cell that
        // was double-clicked (not necessarily the active cell)."
        Check("ON.DOUBLECLICK trap", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!B2");

        // "A command associated with an xlcOnEnter ... | A reference to the
        // cell or cells being entered." CELLS -- so a range, not just a cell.
        op.val.sref.ref.rwLast = 3; op.val.sref.ref.colLast = 2;
        Check("ON.ENTER trap, a range being entered", op, "[Book1]Sheet1", "cell", "[Book1]Sheet1!B2:C4");
    }
    {
        // "Auto_Open, AutoClose, Auto_Activate or Auto_Deactivate macro | The
        // name of the calling sheet." A STRING, so it decodes down the same
        // path as a graphic object -- which is why the trace reads
        // `object:[Book1]Sheet1` for an Auto macro and why that is correct
        // rather than a misclassification.
        auto s = Pascal(L"[Book1]Sheet1");
        XLOPER12 op{}; op.xltype = xltypeStr; op.val.str = s.data();
        Check("Auto_Open: the calling sheet's name", op, nullptr, "name", "[Book1]Sheet1");
    }
    {
        // "DLL | The Register ID." Already covered as `num:` below, but the
        // row deserves its own name.
        XLOPER12 op{}; op.xltype = xltypeNum; op.val.num = 42;
        Check("called from the DLL itself: the Register ID", op, nullptr, "registerid", "42");
    }

    // ---- toolbar and menu ------------------------------------------------
    {
        XLOPER12 two[2]{};
        two[0].xltype = xltypeNum; two[0].val.num = 5;
        two[1].xltype = xltypeInt; two[1].val.w  = 2;
        XLOPER12 op{}; op.xltype = xltypeMulti;
        op.val.array.rows = 1; op.val.array.columns = 2; op.val.array.lparray = two;
        Check("toolbar tool, built-in (numbers)", op, nullptr, "toolbar", "5/2");

        // A custom toolbar answers with its name, not a number, so the name is quoted and never
        // reads like a number.
        auto barName = Pascal(L"XRayProbeBar");
        two[0].xltype = xltypeStr; two[0].val.str = barName.data();
        Check("toolbar tool, custom (named)", op, nullptr, "toolbar", "\"XRayProbeBar\"/2");
        two[0].xltype = xltypeNum; two[0].val.num = 5;

        // A menu is four elements: {bar ID, menu, submenu, command}.
        XLOPER12 four[4]{};
        four[0].xltype = xltypeNum; four[0].val.num = 1;
        four[1].xltype = xltypeNum; four[1].val.num = 2;
        four[2].xltype = xltypeNum; four[2].val.num = 3;
        four[3].xltype = xltypeInt; four[3].val.w   = 4;
        op.val.array.columns = 4; op.val.array.lparray = four;
        Check("menu command", op, nullptr, "menu", "1/2/3/4");

        // Three is NOT a menu, and must not be reported as one.
        op.val.array.columns = 3; op.val.array.lparray = four;
        Check("three elements is not a menu", op, nullptr, "array", "3");

        // A size nobody has documented is named by its size, not guessed at.
        op.val.array.columns = 7; op.val.array.lparray = four;
        Check("array of 7", op, nullptr, "array", "7");

        op.val.array.columns = 2; op.val.array.lparray = nullptr;
        Check("array, null lparray", op, nullptr, "array", "2");
    }

    // No caller on a sheet: what a macro, an event handler and Application.Run return.
    {
        XLOPER12 op{}; op.xltype = xltypeErr;
        op.val.err = xlerrRef;
        Check("macro / event / Run", op, nullptr, "none", "ref");
        op.val.err = xlerrValue;
        Check("some other error", op, nullptr, "none", "value");
        op.val.err = 999;
        Check("an error code we do not know", op, nullptr, "none", "err999");
    }

    // ---- the remaining shapes -------------------------------------------
    {
        XLOPER12 op{}; op.xltype = xltypeNum; op.val.num = 12;
        Check("a bare number", op, nullptr, "registerid", "12");

        XLOPER12 nil{}; nil.xltype = xltypeNil;
        Check("nil", nil, nullptr, "none", "nil");

        XLOPER12 miss{}; miss.xltype = xltypeMissing;
        Check("missing", miss, nullptr, "none", "nil");

        // xltypeBool is 0x4. (0x100 is xltypeNil -- getting that wrong is
        // exactly the sort of thing this harness is for.)
        XLOPER12 boolean{}; boolean.xltype = xltypeBool; boolean.val.xbool = 1;
        Check("a type we have not seen", boolean, nullptr, "unknown", "0x4");
    }

    std::printf("\n  %s\n", g_fail ? "FAILURES ABOVE" : "every caller kind decodes as stated");
    return g_fail ? 1 : 0;
}
