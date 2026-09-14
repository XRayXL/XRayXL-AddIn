#pragma once

// Byte-wise conversion for text known to be ASCII: keywords, level names,
// module and function names. Anything wider is truncated to its low byte.
namespace core
{
    inline int WidenAscii(const char* in, wchar_t* out, int cap)
    {
        if (cap <= 0) return 0;
        int k = 0;
        for (; in && in[k] && k < cap - 1; ++k)
            out[k] = static_cast<wchar_t>(static_cast<unsigned char>(in[k]));
        out[k] = 0;
        return k;
    }

    inline int NarrowAscii(const wchar_t* in, char* out, int cap)
    {
        if (cap <= 0) return 0;
        int k = 0;
        for (; in && in[k] && k < cap - 1; ++k)
            out[k] = static_cast<char>(static_cast<unsigned char>(in[k]));
        out[k] = 0;
        return k;
    }
}
