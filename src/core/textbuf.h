#pragma once
#include <cstddef>
#include <cstdlib>
#include <cstring>

namespace core
{
    // The most one rendered value may take. An array past it is written as its shape alone
    // (see TextValueWriter), so a buffer never holds part of a value.
    constexpr std::size_t kMaxValueBytes = 4u * 1024 * 1024;

    // What a buffer keeps between values; a larger one is freed by Clear.
    constexpr std::size_t kKeepBytes = 256u * 1024;

    // A NUL-terminated text buffer that grows on demand up to `limit`. A write that would pass
    // the limit is dropped and remembered in Over(), for the owner to refuse what did not fit.
    // Trivially constructible so it can be __declspec(thread); Release() frees it.
    struct TextBuf
    {
        char*       p     = nullptr;
        std::size_t len   = 0;
        std::size_t cap   = 0;
        std::size_t limit = kMaxValueBytes;
        bool        over  = false;
        // Arrays written into this buffer as their shape alone, since the last Clear.
        int         refused = 0;

        const char* Text() const { return p ? p : ""; }
        std::size_t Len()  const { return len; }
        bool        Over() const { return over; }

        void Clear()
        {
            if (cap > kKeepBytes) Release();
            len = 0; over = false; refused = 0;
            if (p) p[0] = 0;
        }

        // Back to `n` bytes, which also forgets an overflow past them.
        void Truncate(std::size_t n) { if (n < len) len = n; if (p) p[len] = 0; over = false; }

        void Release() { free(p); p = nullptr; len = cap = 0; over = false; refused = 0; }

        // Room for `n` more bytes and the NUL, or nullptr once the limit would be passed.
        char* Reserve(std::size_t n)
        {
            if (over) return nullptr;
            const std::size_t need = len + n + 1;
            if (need > limit) { over = true; return nullptr; }
            if (need > cap)
            {
                std::size_t c = cap ? cap : 4096;
                while (c < need) c *= 2;
                if (c > limit) c = limit;
                char* q = static_cast<char*>(realloc(p, c));
                if (!q) { over = true; return nullptr; }
                p = q; cap = c;
            }
            return p + len;
        }

        // After Reserve: `n` bytes were written at the end.
        void Commit(std::size_t n) { len += n; p[len] = 0; }

        void Append(const char* s, std::size_t n)
        {
            char* d = Reserve(n);
            if (!d) return;
            memcpy(d, s, n);
            Commit(n);
        }
        void Append(const char* s) { Append(s, strlen(s)); }
        void Append(char c) { Append(&c, 1); }
    };
}
