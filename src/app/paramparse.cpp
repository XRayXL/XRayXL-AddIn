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

    // ReadUpper truncates at cap-1, so a name at that ceiling would match every longer input
    // starting with it; 32 keeps BUFFERWHENFULL (14) well clear.
    bool IsWord(LPXLOPER12 v, const wchar_t* w)
    {
        if (ArgType(v) != xltypeStr) return false;
        wchar_t b[32]; ReadUpper(v, b, 32);
        return wcscmp(b, w) == 0;
    }

    static constexpr double kBufMaxBytes = 240.0 * 1024.0 * 1024.0;
    bool ParseBufferText(const wchar_t* b, unsigned long long& outBytes);
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
            return ParseBufferText(b, outBytes);
        }
        else return false;
        if (num < 0) return false;
        const double bytes0 = num * static_cast<double>(unit);
        if (bytes0 > kBufMaxBytes) return false;
        outBytes = static_cast<unsigned long long>(bytes0);
        return true;
    }

    // The text half, so the Options dialog and SetTraceParam accept the same words. Upper-cased by the caller.
    bool ParseBufferText(const wchar_t* b, unsigned long long& outBytes)
    {
        if (!b) return false;
        double num = 0.0;
        unsigned long long unit = 1024ull * 1024ull;
        {
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
        // One vocabulary per setting, and no synonyms: a word matches or it does not.
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

    bool ParseFormat(LPXLOPER12 v, core::modes::Format& f)
    {
        static const Word<core::modes::Format> k[] = { { L"CSV",   core::modes::Format::Csv   },
                                                       { L"JSONL", core::modes::Format::Jsonl } };
        return ParseWord(v, k, f);
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
            { L"OBJECTS", core::modes::Param::Objects }, { L"BREAKPOINTS", core::modes::Param::Breakpoints } };
        return ParseWord(v, k, p);
    }

    // A native boolean and the text "TRUE" are the same word: Excel delivers either, depending on
    // whether it came from a cell, a formula or Application.Run.
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
