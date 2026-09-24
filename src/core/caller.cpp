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
        // Column letters are bijective base-26, with no zero digit: 26 is Z and 27 is AA, which a
        // plain base-26 conversion gets wrong.
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

        // Not named `pascal`: windef.h defines that as a calling convention.
        void Utf8From(const XCHAR* pstr, char* out, int outSize)
        {
            out[0] = 0;
            if (!pstr || pstr[0] <= 0) return;
            NarrowUtf8(pstr + 1, pstr[0], out, outSize);
        }

        // "none:ref" is the macro dialog, an Auto macro, an event or the VBE; "none:err<N>" is
        // something else and must not be filed under the same heading.
        const char* ErrName(int e) { return ExcelErrShortName(e); }

        double NumOf(const XLOPER12& x)
        {
            const int t = x.xltype & kXlTypeMask;
            if (t == xltypeNum) return x.val.num;
            if (t == xltypeInt) return static_cast<double>(x.val.w);
            return 0.0;
        }

        // Excel's own quoting rule:
        //
        //    [Plain1.xlsx]Sheet1!A1        a dot alone does not quote
        //    [Under_score.xlsx]Under_1!A1  underscore is safe
        //    [Plain5.xlsx]A.B!A1           a dot in the sheet is safe too
        //    '[has-hyphen.xlsx]Sheet1'!A1  a hyphen quotes, either side
        //    '[has space.xlsx]Sheet1'!A1   so does a space
        //    '[Digits123.xlsx]1Sheet'!A1   and a sheet name starting with a digit
        //    '[Plain4.xlsx]Bob''s'!A1      an apostrophe is doubled inside
        //
        // The safe set is no wider than this. Quoting too much still pastes back; quoting too
        // little does not.
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

        // A custom toolbar puts its name where a built-in one puts a number, so a string is quoted
        // and never reads as a number.
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
        // Both corners: a CSE array formula across B2:D4 is one formula and xlfCaller names all
        // nine cells, so the first corner alone would name a cell no formula occupies.
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
            // reftbl[0] only: Excel refuses to enter a multi-area array formula, and the first area
            // would still be true.
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
            // A graphic object's name: the only attribution the call has, and never a cell.
            char name[512];
            Utf8From(caller.val.str, name, sizeof(name));
            // Empty from a non-empty name means it did not fit, not that there is none.
            const bool tooLong = (caller.val.str != nullptr && caller.val.str[0] > 0 && name[0] == 0);
            Say(out, "name", "%s", name[0] ? name : (tooLong ? "nametoolong" : "(unnamed)"));
            break;
        }

        case xltypeMulti:
        {
            // Two elements is a toolbar tool {toolbar, position}; four is a menu command {bar ID,
            // menu, submenu, command}. Any other shape is written as its size.
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
            // #REF! is the documented answer for "not called from a sheet": an answer, not a failure.
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
            // A zero-count SRef/Ref reaches here having set nothing, and `kind` is never empty.
            if (!out.kind[0]) Say(out, "none", "emptyref");
            return;
        }

        // "B2", or "B2:D4" when the caller spans more than one cell. The widest is
        // "XFC1048575:XFD1048576", 21 characters, so ref[32] holds any range.
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

        // Excel's own calls arrive as row 0 column 0 with no sheet, which decodes to a confident
        // "A1", so the cell is written only when the sheet resolved.
        if (sheetName && sheetName[0])
        {
            // Quoted exactly where Excel quotes it, so the text pastes back.
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

        // Never xlCoerce: it fails with xlretUncalced on an uncalculated cell.
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
                    // WideCharToMultiByte fails rather than truncates, so a name too long to
                    // convert would otherwise read as a sheetless caller.
                    nameDidNotFit = (nm.val.str != nullptr && nm.val.str[0] > 0 && sheet[0] == 0);
                }
                Excel12(xlFree, nullptr, 1, &nm);
            }
        }

        DecodeCaller(&caller, sheet, out);
        Excel12(xlFree, nullptr, 1, &caller);

        // After the decode, so DecodeCaller stays pure: it is handed a sheet name or not, never why.
        if (nameDidNotFit && !out.isCell) Say(out, "none", "nametoolong");
    }
}
