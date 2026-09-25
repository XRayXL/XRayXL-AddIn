#include "sheetquote.h"

#include <cstdio>
#include <cstring>

namespace core
{
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
}
