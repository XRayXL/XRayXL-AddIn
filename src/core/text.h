#pragma once
#include <windows.h>
#include <cstring>

namespace core
{
    // UTF-8 from UTF-16 into a fixed buffer, always NUL-terminated. `chars` is
    // the source length, or -1 for a NUL-terminated source. Returns the bytes
    // written; 0, with an empty buffer, when nothing fit or nothing was there.
    // What did not fit is dropped whole rather than cut mid-character.
    inline int NarrowUtf8(const wchar_t* w, int chars, char* dst, int cap)
    {
        if (cap <= 0) return 0;
        dst[0] = 0;
        if (w == nullptr || chars == 0) return 0;
        int n = WideCharToMultiByte(CP_UTF8, 0, w, chars, dst, cap - 1, nullptr, nullptr);
        if (n <= 0) { dst[0] = 0; return 0; }
        if (chars < 0 && dst[n - 1] == 0) n--;      // a -1 source counts its own NUL
        dst[n] = 0;
        return n;
    }

    // UTF-16 from UTF-8 into a fixed buffer, always NUL-terminated. A source too
    // long is cut at a character boundary. Returns the characters written.
    inline int WidenUtf8(const char* s, wchar_t* dst, int cap)
    {
        if (cap <= 0) return 0;
        dst[0] = 0;
        if (s == nullptr || *s == 0) return 0;
        int len = static_cast<int>(strlen(s));
        if (len > cap - 1)
        {
            len = cap - 1;
            while (len > 0 && (static_cast<unsigned char>(s[len]) & 0xC0) == 0x80) --len;
        }
        int n = MultiByteToWideChar(CP_UTF8, 0, s, len, dst, cap - 1);
        if (n < 0) n = 0;
        dst[n] = 0;
        return n;
    }
}
