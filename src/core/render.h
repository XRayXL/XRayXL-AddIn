#pragma once
#include <cstdio>
#include <cstring>

// WHAT BOTH COLUMNS AGREE ON. The XLL and VBA rows share one trace file, so a
// number and a capped array must read the same whichever source wrote them.
namespace core
{
    // A cap on rendered array elements, so a multi-million element array cannot spin. The real
    // limit is the output buffer, and the true count is always in the header.
    constexpr int kMaxRenderedElems = 1 << 20;

    // %.15g: fifteen digits round-trip every value a double holds exactly;
    // %.17g adds only noise digits. Two renderings of one number that disagree
    // are worse than either.
    inline void FormatDouble(double d, char* out, int cap)
    {
        _snprintf_s(out, cap, _TRUNCATE, "%.15g", d);
    }

    // Appends `s` at `len`, clipping at `cap`. Returns the new length.
    inline int Append(char* buf, int cap, int len, const char* s)
    {
        if (len < 0 || len >= cap - 1) return len;
        const int n = static_cast<int>(strlen(s));
        const int room = cap - 1 - len;
        const int take = n < room ? n : room;
        memcpy(buf + len, s, static_cast<size_t>(take));
        buf[len + take] = 0;
        return len + take;
    }

    // The marker for a list cut short, in the same words for both columns.
    inline int AppendShownMarker(char* buf, int cap, int len,
                                 unsigned long long shown, unsigned long long total)
    {
        char m[64];
        _snprintf_s(m, sizeof(m), _TRUNCATE, ",...(%llu of %llu shown)", shown, total);
        return Append(buf, cap, len, m);
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
