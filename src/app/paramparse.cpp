#include "paramparse.h"
#include "core/ascii.h"
#include "core/render.h"
#include <cwctype>
#include <cstdlib>
#include <cwchar>
#include <cstring>

namespace app
{

namespace params
{
    void ReadUpper(const XLOPER12* v, wchar_t* buf, int cap)
    {
        const XCHAR* p = v->val.str;
        int n = p ? p[0] : 0;
        if (n > cap - 1) n = cap - 1;
        for (int i = 0; i < n; i++) buf[i] = static_cast<wchar_t>(towupper(p[1 + i]));
        buf[n] = 0;
    }

    int ArgType(LPXLOPER12 v)
    {
        return v ? static_cast<int>(v->xltype & ~(xlbitXLFree | xlbitDLLFree))
                 : xltypeMissing;
    }
    bool IsMissing(LPXLOPER12 v)
    {
        const int t = ArgType(v);
        return t == xltypeMissing || t == xltypeNil;
    }

    // ReadUpper TRUNCATES at cap-1, so the buffer sets a ceiling on how long a
    // setting name may be; a name AT the ceiling compares equal to every longer
    // input starting with it. 32/31 keeps BUFFERWHENFULL (14) well clear.
    bool IsWord(LPXLOPER12 v, const wchar_t* w)
    {
        if (ArgType(v) != xltypeStr) return false;
        wchar_t b[32]; ReadUpper(v, b, 32);
        return wcscmp(b, w) == 0;
    }

    // THE OUTPUT BUFFER SIZE, in bytes, from a number with an optional unit.
    // A bare number, or M/MB, is megabytes; K/KB is kilobytes -- so
    // "10", "10M", "10MB" are 10 MB and "10K"/"10KB" are 10 KB, case-
    // insensitive. 0 is synchronous. A native number argument (no text, no unit)
    // is megabytes. Returns false if the text is not a number with an accepted
    // unit; the caller enforces the ring floor and the 4 GB cap.
    static constexpr double kBufMaxBytes = 4096.0 * 1024.0 * 1024.0;   // 4 GB
    bool ParseBufferBytes(LPXLOPER12 v, unsigned long long& outBytes)
    {
        double num = 0.0;
        unsigned long long unit = 1024ull * 1024ull;           // default: megabytes
        const int t = ArgType(v);
        if (t == xltypeInt) { num = v->val.w; }
        else if (t == xltypeNum) { num = v->val.num; }
        else if (t == xltypeStr)
        {
            wchar_t b[24]; ReadUpper(v, b, 24);
            int i = 0; bool dot = false, any = false;
            wchar_t numbuf[24]; int nb = 0;
            while (b[i] && nb < 23 && ((b[i] >= L'0' && b[i] <= L'9') || (b[i] == L'.' && !dot)))
            { if (b[i] == L'.') dot = true; else any = true; numbuf[nb++] = b[i++]; }
            numbuf[nb] = 0;
            if (!any) return false;                            // no digits: a lone dot is not a number
            while (b[i] == L' ') i++;                           // "10 KB" as well as "10KB"
            const wchar_t* u = b + i;
            if      (!u[0] || !wcscmp(u, L"M") || !wcscmp(u, L"MB")) unit = 1024ull * 1024ull;
            else if (!wcscmp(u, L"K") || !wcscmp(u, L"KB"))          unit = 1024ull;
            else return false;                                 // unknown unit
            // Not _wtof: under a comma-decimal locale it stops "1.5" at the dot.
            double whole = 0.0, frac = 0.0, scale = 1.0;
            bool seenDot = false;
            for (int k = 0; k < nb; ++k)
            {
                if (numbuf[k] == L'.') { seenDot = true; continue; }
                const double d = static_cast<double>(numbuf[k] - L'0');
                if (!seenDot) whole = whole * 10.0 + d;
                else          { scale *= 10.0; frac += d / scale; }
            }
            num = whole + frac;
        }
        else return false;
        if (num < 0) return false;
        const double bytes = num * static_cast<double>(unit);
        if (bytes > kBufMaxBytes) return false;
        outBytes = static_cast<unsigned long long>(bytes);
        return true;
    }

    void FormatBufferW(unsigned long long bytes, wchar_t* out, int cap)
    {
        // FormatBytes without its space, so ParseBufferBytes reads it back.
        char s[32] = "0";
        if (bytes) core::FormatBytes(bytes, s, sizeof(s));
        char t[32];
        int k = 0;
        for (int i = 0; s[i] && k < 31; ++i) if (s[i] != ' ') t[k++] = s[i];
        t[k] = 0;
        core::WidenAscii(t, out, cap);
    }

    namespace
    {
        // ONE VOCABULARY PER SETTING, and no synonyms: a word matches or it
        // does not.
        template <class T> struct Word { const wchar_t* text; T value; };

        template <class T, size_t N>
        bool ParseWord(LPXLOPER12 v, const Word<T> (&words)[N], T& out)
        {
            if (ArgType(v) != xltypeStr) return false;
            wchar_t b[16]; ReadUpper(v, b, 16);
            for (const Word<T>& w : words)
                if (!wcscmp(b, w.text)) { out = w.value; return true; }
            return false;
        }
    }

    bool ParseWhenFull(LPXLOPER12 v, bool& pause)
    {
        static const Word<bool> k[] = { { L"DROP", false }, { L"PAUSE", true } };
        return ParseWord(v, k, pause);
    }

    bool ParseSource(LPXLOPER12 v, bool& both, core::modes::Source& s)
    {
        if (IsMissing(v)) { both = true; return true; }
        both = false;
        static const Word<core::modes::Source> k[] = { { L"XLL", core::modes::Source::Xll },
                                                 { L"VBA", core::modes::Source::Vba } };
        return ParseWord(v, k, s);
    }

    bool ParseParam(LPXLOPER12 v, core::modes::Param& p)
    {
        static const Word<core::modes::Param> k[] = {
            { L"DEPTH",  core::modes::Param::Depth  }, { L"ARGS",   core::modes::Param::Args   },
            { L"RETVAL", core::modes::Param::RetVal },
            { L"OBJECTS", core::modes::Param::Objects } };
        return ParseWord(v, k, p);
    }

    // TRUE or FALSE. A native boolean and the text "TRUE" are the SAME word in
    // two encodings -- which is how Excel delivers it, depending on whether it
    // came from a cell, a formula or Application.Run -- so both are read.
    bool ParseOnOff(LPXLOPER12 v, bool& on)
    {
        if (ArgType(v) == xltypeBool) { on = v->val.xbool != 0; return true; }
        static const Word<bool> k[] = { { L"TRUE", true }, { L"FALSE", false } };
        return ParseWord(v, k, on);
    }

    bool ParseDepth(LPXLOPER12 v, core::modes::Depth& d)
    {
        static const Word<core::modes::Depth> k[] = { { L"OFF", core::modes::Depth::Off },
                                                { L"TOP", core::modes::Depth::Top },
                                                { L"ALL", core::modes::Depth::All } };
        return ParseWord(v, k, d);
    }
}
}   // namespace app
