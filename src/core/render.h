#pragma once
#include <cstdio>
#include <cstring>

// The XLL and VBA rows share one trace file, so a number must read the same whichever source
// wrote it.
namespace core
{
    // %.15g, not %.17g: the extra digits are noise, and two renderings of one number that disagree
    // are worse than either.
    inline void FormatDouble(double d, char* out, int cap)
    {
        _snprintf_s(out, cap, _TRUNCATE, "%.15g", d);
    }

    // "1 MB" or "512 KB": whole megabytes when the size divides evenly.
    inline void FormatBytes(unsigned long long bytes, char* out, int cap)
    {
        const unsigned long long mb = 1024ull * 1024ull;
        if (bytes % mb == 0) _snprintf_s(out, cap, _TRUNCATE, "%llu MB", bytes / mb);
        else                 _snprintf_s(out, cap, _TRUNCATE, "%llu KB", bytes / 1024ull);
    }

    // A string quoted and escaped so it survives its CSV field unchanged:
    // \" \\ \t \r \n, \xNN below 32 and DEL, \uNNNN above 126. `take`
    // characters are rendered; `cut` says the source was longer. Truncation
    // lands on a character boundary and ends "...".
    inline bool RenderQuoted(const wchar_t* w, int take, bool cut, char* out, int cap)
    {
        if (cap < 16) return false;
        static const char kHex[] = "0123456789ABCDEF";
        constexpr int kTail = 4;              // the closing quote and "..."
        const int lim = cap - 1 - kTail;
        int j = 0;
        out[j++] = '"';
        for (int i = 0; i < take; ++i)
        {
            const wchar_t c = w[i];
            char seq[8];
            int  len = 0;
            if      (c == L'\t') { seq[len++] = '\\'; seq[len++] = 't'; }
            else if (c == L'\r') { seq[len++] = '\\'; seq[len++] = 'r'; }
            else if (c == L'\n') { seq[len++] = '\\'; seq[len++] = 'n'; }
            else if (c == L'\\') { seq[len++] = '\\'; seq[len++] = '\\'; }
            else if (c == L'\"') { seq[len++] = '\\'; seq[len++] = '\"'; }
            else if (c < 32 || c == 127)
            {
                seq[len++] = '\\'; seq[len++] = 'x';
                seq[len++] = kHex[(c >> 4) & 0xF]; seq[len++] = kHex[c & 0xF];
            }
            else if (c > 126)
            {
                seq[len++] = '\\'; seq[len++] = 'u';
                seq[len++] = kHex[(c >> 12) & 0xF]; seq[len++] = kHex[(c >> 8) & 0xF];
                seq[len++] = kHex[(c >>  4) & 0xF]; seq[len++] = kHex[c & 0xF];
            }
            else seq[len++] = static_cast<char>(c);

            if (j + len > lim) { cut = true; break; }
            for (int k = 0; k < len; ++k) out[j++] = seq[k];
        }
        out[j++] = '"';
        if (cut) { out[j++] = '.'; out[j++] = '.'; out[j++] = '.'; }
        out[j] = 0;
        return true;
    }
}
