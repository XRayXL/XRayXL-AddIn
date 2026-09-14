#include "caller.h"
#include "excel_api.h"
#include "text.h"
#include "excelerr.h"
#include "xltype.h"

#include <cstdio>
#include <cstring>
#include <cstdarg>

namespace core
{
    namespace
    {
        // "A1" from a zero-based row and column. Excel's column letters are
        // bijective base-26 -- there is no zero digit -- so 26 is Z and 27 is
        // AA, which a plain base-26 conversion gets wrong.
        void RefText(int row0, int col0, char* out, int outSize)
        {
            char rev[8]; int n = 0;
            int c = col0 + 1;
            while (c > 0 && n < 6) { rev[n++] = static_cast<char>('A' + (c - 1) % 26); c = (c - 1) / 26; }
            char col[8];
            for (int i = 0; i < n; ++i) col[i] = rev[n - 1 - i];
            col[n] = 0;
            _snprintf_s(out, outSize, _TRUNCATE, "%s%d", col, row0 + 1);
        }

        // str[0] is the length, characters follow, no null terminator.
        // (`pascal` is taken: windef.h defines it as a calling convention.)
        void Utf8From(const XCHAR* pstr, char* out, int outSize)
        {
            out[0] = 0;
            if (!pstr || pstr[0] <= 0) return;
            NarrowUtf8(pstr + 1, pstr[0], out, outSize);
        }

        // "none:ref" says the caller was the macro dialog, an Auto macro, an
        // event or the VBE. "none:err<N>" is something else entirely and should
        // be looked at rather than filed under the same heading.
        const char* ErrName(int e) { return ExcelErrShortName(e); }

        double NumOf(const XLOPER12& x)
        {
            const int t = x.xltype & kXlTypeMask;
            if (t == xltypeNum) return x.val.num;
            if (t == xltypeInt) return static_cast<double>(x.val.w);
            return 0.0;
        }

        // ONE ELEMENT OF A TOOLBAR OR MENU ARRAY, BY ITS OWN TYPE.
        //
        // These were rendered with NumOf and "%g", which returns 0.0 for
        // anything that is not a number -- silently. The SDK says a toolbar
        // answer is "the toolbar number for built-in toolbars OR THE TOOLBAR
        // NAME for custom toolbars", so a custom bar puts a STRING in that slot
        // and it printed as `0`: the name lost, and indistinguishable from a
        // real zero. A reader could not tell `toolbar:1/0` with a numeric 1
        // from the same text produced by flattening a string.
        //
        // Quoted when it is a string, so the two can never read alike.

        // EXCEL'S OWN QUOTING RULE, measured rather than recalled.
        //
        //   [Plain1.xlsx]Sheet1!A1        a dot alone does NOT quote -- so an
        //                                 ordinary saved workbook is bare
        //   [Under_score.xlsx]Under_1!A1  underscore is safe
        //   [Plain5.xlsx]A.B!A1           a dot in the SHEET is safe too
        //   '[has-hyphen.xlsx]Sheet1'!A1  a hyphen quotes, either side
        //   '[has space.xlsx]Sheet1'!A1   so does a space
        //   '[Digits123.xlsx]1Sheet'!A1   and a SHEET NAME STARTING WITH A DIGIT
        //   '[Plain4.xlsx]Bob''s'!A1  an apostrophe is DOUBLED inside
        //
        // [measured: a planted-name probe over each of those shapes]
        //
        // The safe set is deliberately the one that was MEASURED and no wider.
        // Quoting where Excel would not is a cosmetic difference in a reference
        // that still pastes back; failing to quote where Excel would produces one
        // that does not, so the doubt falls on the side of the quote.
        void QuoteSheetPrefix(const char* prefix, char* out, int cap)
        {
            const char* sheet = std::strchr(prefix, ']');
            sheet = sheet ? sheet + 1 : prefix;

            bool needs = (*sheet >= '0' && *sheet <= '9');
            for (const char* p = prefix; *p && !needs; ++p)
            {
                const char c = *p;
                if (c == '[' || c == ']') continue;
                const bool safe = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
                                  (c >= '0' && c <= '9') || c == '_' || c == '.';
                if (!safe) needs = true;
            }
            if (!needs) { _snprintf_s(out, cap, _TRUNCATE, "%s", prefix); return; }

            int j = 0;
            const int lim = cap - 2;               // the closing quote and the NUL
            if (j < lim) out[j++] = '\'';
            for (const char* p = prefix; *p && j < lim; ++p)
            {
                if (*p == '\'' && j < lim - 1) out[j++] = '\'';   // doubled, as Excel does
                out[j++] = *p;
            }
            if (j < cap - 1) out[j++] = '\'';
            out[j] = 0;
        }

        void ElemText(const XLOPER12& x, char* out, int cap)
        {
            const int t = x.xltype & kXlTypeMask;
            switch (t)
            {
            case xltypeNum:  _snprintf_s(out, cap, _TRUNCATE, "%g", x.val.num);            break;
            case xltypeInt:  _snprintf_s(out, cap, _TRUNCATE, "%d", static_cast<int>(x.val.w)); break;
            case xltypeBool: _snprintf_s(out, cap, _TRUNCATE, "%s", x.val.xbool ? "TRUE" : "FALSE"); break;
            case xltypeStr:
            {
                char s[192];
                Utf8From(x.val.str, s, sizeof(s));
                _snprintf_s(out, cap, _TRUNCATE, "\"%s\"", s);
                break;
            }
            case xltypeErr:
            {
                const char* e = ErrName(x.val.err);
                if (e) _snprintf_s(out, cap, _TRUNCATE, "err:%s", e);
                else   _snprintf_s(out, cap, _TRUNCATE, "err:%d", x.val.err);
                break;
            }
            case xltypeNil:
            case xltypeMissing: _snprintf_s(out, cap, _TRUNCATE, "-");  break;
            // An unknown shape is named, never rendered as a plausible value.
            default: _snprintf_s(out, cap, _TRUNCATE, "type:0x%X", t);  break;
            }
        }
    }

    namespace
    {
        void Say(Caller& out, const char* kind, const char* fmt = nullptr, ...)
        {
            _snprintf_s(out.kind, _TRUNCATE, "%s", kind);
            if (!fmt) { out.desc[0] = 0; return; }
            va_list ap; va_start(ap, fmt);
            _vsnprintf_s(out.desc, sizeof(out.desc), _TRUNCATE, fmt, ap);
            va_end(ap);
        }
    }

    void DecodeCaller(const void* callerOper, const char* sheetName, Caller& out)
    {
        out.kind[0] = 0;
        out.desc[0] = 0;
        out.isCell = false;

        const XLOPER12& caller = *static_cast<const XLOPER12*>(callerOper);
        const int t = caller.xltype & kXlTypeMask;
        // BOTH CORNERS. A legacy CSE array formula entered across B2:D4 is ONE
        // formula occupying nine cells and xlfCaller names all of them. Reading
        // only rwFirst/colFirst wrote that as "B2" -- a cell no formula occupies
        // on its own, indistinguishable from a real single-cell caller, and
        // wrong in the direction that reads as a fact.
        int r0 = -1, c0 = -1, r1 = -1, c1 = -1;

        switch (t)
        {
        case xltypeSRef:
            if (caller.val.sref.count >= 1)
            {
                r0 = caller.val.sref.ref.rwFirst;
                c0 = caller.val.sref.ref.colFirst;
                r1 = caller.val.sref.ref.rwLast;
                c1 = caller.val.sref.ref.colLast;
            }
            break;

        case xltypeRef:
            // reftbl[0] only. A caller with more than one area would need a
            // multi-area array formula, which Excel refuses to enter; if one
            // ever arrives, the first area is a true statement about where the
            // formula is rather than a guess.
            if (caller.val.mref.lpmref != nullptr && caller.val.mref.lpmref->count >= 1)
            {
                r0 = caller.val.mref.lpmref->reftbl[0].rwFirst;
                c0 = caller.val.mref.lpmref->reftbl[0].colFirst;
                r1 = caller.val.mref.lpmref->reftbl[0].rwLast;
                c1 = caller.val.mref.lpmref->reftbl[0].colLast;
            }
            break;

        case xltypeStr:
        {
            // A macro run from a graphic object. The string is that object's
            // NAME and the only attribution such a call has -- there is no
            // cell, and putting the name in the cell column would be a
            // confident wrong one.
            char name[512];
            Utf8From(caller.val.str, name, sizeof(name));
            // Empty from a non-empty name means it did not fit, not that there is none.
            const bool tooLong = (caller.val.str != nullptr && caller.val.str[0] > 0 && name[0] == 0);
            Say(out, "name", "%s", name[0] ? name : (tooLong ? "nametoolong" : "(unnamed)"));
            break;
        }

        case xltypeMulti:
        {
            // TWO elements is a toolbar tool {toolbar, position}; FOUR is a menu
            // command {bar ID, menu, submenu, command}. Anything else is a shape
            // this decoder has not seen, and is written as its size rather than
            // forced into one of the two.
            // [published: XLL SDK, xlfCaller]
            const int n = caller.val.array.rows * caller.val.array.columns;
            const XLOPER12* a = caller.val.array.lparray;
            if (a && (n == 2 || n == 4))
            {
                char e[4][208];
                for (int i = 0; i < n; ++i) ElemText(a[i], e[i], sizeof(e[i]));
                if (n == 2) Say(out, "toolbar", "%s/%s", e[0], e[1]);
                else        Say(out, "menu", "%s/%s/%s/%s", e[0], e[1], e[2], e[3]);
            }
            else
                Say(out, "array", "%d", n);
            break;
        }

        case xltypeErr:
        {
            // #REF! is the documented answer for "not called from a sheet" --
            // the macro dialog, an Auto macro, an event handler, the VBE,
            // Application.Run, DDE/OLE. An ANSWER, not a failure.
            const char* e = ErrName(caller.val.err);
            if (e) Say(out, "none", "%s", e);
            else   Say(out, "none", "err%d", caller.val.err);
            break;
        }

        case xltypeNum:
        case xltypeInt:
            Say(out, "registerid", "%g", NumOf(caller));
            break;

        case xltypeNil:
        case xltypeMissing:
            Say(out, "none", "nil");
            break;

        default:
            // Named by its type bits rather than swallowed: a caller kind
            // nobody has seen is worth one line in a trace, not silence.
            Say(out, "unknown", "0x%X", t);
            break;
        }

        if (r0 < 0 || c0 < 0)
        {
            // A REFERENCE THAT NAMES NOTHING: a zero-count SRef/Ref reaches
            // here having set nothing, since the switch above fills `what` only
            // for the non-reference kinds. caller.h promises `what` is never
            // empty, and an empty one reads in the trace exactly like "we never
            // asked". [measured]
            if (!out.kind[0]) Say(out, "none", "emptyref");
            return;
        }

        // "B2", or "B2:D4" when the caller spans more than one cell. The widest
        // this can be is two far-corner refs and a colon -- "XFC1048575:XFD1048576",
        // 21 characters -- so ref[32] holds any range Excel can produce. Refs are
        // ASCII, so characters and bytes are the same count here; the SHEET name
        // is the field where they are not.
        char ref[32];
        RefText(r0, c0, ref, sizeof(ref));
        if (r1 > r0 || c1 > c0)
        {
            char last[16];
            RefText(r1 < r0 ? r0 : r1, c1 < c0 ? c0 : c1, last, sizeof(last));
            char span[32];
            _snprintf_s(span, _TRUNCATE, "%s:%s", ref, last);
            _snprintf_s(ref, _TRUNCATE, "%s", span);
        }

        // THE ROW-0 COLUMN-0 TRAP (caller.h): Excel calls functions for its
        // own purposes with no calling cell, and those arrive as row 0 column 0
        // with no sheet, decoding arithmetically to a confident "A1". A real
        // calculation names a real sheet, so the cell is written only when the
        // sheet resolved.
        if (sheetName && sheetName[0])
        {
            // ONE FIELD: "[Book1]Sheet1!B2", or "…!B2:D4" for a CSE range, and
            // QUOTED EXACTLY WHERE EXCEL QUOTES IT.
            char pre[512];
            QuoteSheetPrefix(sheetName, pre, sizeof(pre));
            Say(out, "cell", "%s!%s", pre, ref);
            out.isCell = true;
        }
        else
        {
            Say(out, "none", "sheetless-%s", ref);
        }
    }

    void ReadCaller(Caller& out)
    {
        out.kind[0] = 0;
        out.desc[0] = 0;
        out.isCell = false;

        XLOPER12 caller{};
        const int rc = Excel12(xlfCaller, &caller, 0);
        if (rc != xlretSuccess)
        {
            // Excel declining is a fact about the context we asked from, not
            // an absent caller: the two must not print the same thing.
            Say(out, "unavailable", "%d", rc);
            return;
        }

        // Asked for separately, and only the reference kinds have one.
        // xlCoerce is never used: it fails with xlretUncalced on an
        // uncalculated cell.
        char sheet[512] = { 0 };
        bool nameDidNotFit = false;
        const int t = caller.xltype & kXlTypeMask;
        if (t == xltypeSRef || t == xltypeRef)
        {
            XLOPER12 nm{};
            if (Excel12(xlSheetNm, &nm, 1, &caller) == xlretSuccess)
            {
                if ((nm.xltype & kXlTypeMask) == xltypeStr)
                {
                    Utf8From(nm.val.str, sheet, sizeof(sheet));
                    // "Excel gave us a name we could not hold" is not the same
                    // fact as "there is no sheet", and both arrive here as an
                    // empty string. WideCharToMultiByte does not truncate on an
                    // undersized buffer -- it fails -- so a name too long to
                    // convert would otherwise be reported as a sheetless
                    // caller, which is a false negative about a real cell.
                    nameDidNotFit = (nm.val.str != nullptr && nm.val.str[0] > 0 && sheet[0] == 0);
                }
                Excel12(xlFree, nullptr, 1, &nm);
            }
        }

        DecodeCaller(&caller, sheet, out);
        Excel12(xlFree, nullptr, 1, &caller);

        // Said AFTER the decode so DecodeCaller stays pure and offline-testable:
        // it is handed a sheet name or not, and knows nothing of why.
        if (nameDidNotFit && !out.isCell) Say(out, "none", "nametoolong");
    }
}
