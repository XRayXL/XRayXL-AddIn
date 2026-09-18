// The guarded readers. Every byte read out of Excel or VBE7 is somebody else's, so a stale pointer
// must cost the read and nothing more.
//
// Header-inline leaf functions, because a __try cannot share a frame with C++ unwinding. The
// predicates check an address range and an alignment: they permit a read and say nothing about the
// type there.
#pragma once
#include <windows.h>
#include <cstdint>
#include <cstring>

namespace core
{
    // Below 0x10000 is the reserved null page, where a small integer mistaken
    // for a pointer usually lands. The ceiling is the top of the x64 user half.
    inline bool InUserRange(std::uint64_t p)
    {
        return p >= 0x10000ull && p < 0x00007FFFFFFFFFFFull;
    }

    // 8 for a structure, 4 for a p-code trailer,
    // 2 for a BSTR -- it points four bytes past its length prefix. Unaligned is
    // a reason to REFUSE the read, not a reason to believe anything about what
    // is there.
    inline bool Aligned(std::uint64_t p, std::uint64_t n) { return (p & (n - 1)) == 0; }

    // Named for its conjuncts and nothing else: `PlausibleStruct(p)` read as a
    // verdict on what p points at, which is not a question this can answer.
    inline bool InRangeAndAligned(std::uint64_t p, std::uint64_t n)
    {
        return InUserRange(p) && Aligned(p, n);
    }

    // Each returns false and leaves `out` untouched on a fault. None validates
    // the address: relying on the guard is legitimate -- refusing to read is
    // not the same as being unable to.
    inline bool RdU64(std::uint64_t at, std::uint64_t& out)
    {
        __try { out = *reinterpret_cast<const std::uint64_t*>(at); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    inline bool RdU32(std::uint64_t at, std::uint32_t& out)
    {
        __try { out = *reinterpret_cast<const std::uint32_t*>(at); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    inline bool RdI32(std::uint64_t at, std::int32_t& out)
    {
        __try { out = *reinterpret_cast<const std::int32_t*>(at); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    inline bool RdU16(std::uint64_t at, std::uint16_t& out)
    {
        __try { out = *reinterpret_cast<const std::uint16_t*>(at); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    // A Byte array's last element is one byte, and RdU16 would read past the
    // end of the allocation -- a refusal that depends on where the array sits
    // in a page is not a fact about the program.
    inline bool RdU8(std::uint64_t at, std::uint8_t& out)
    {
        __try { out = *reinterpret_cast<const std::uint8_t*>(at); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    // `n` bytes, for a structure too wide for one register read -- an XLOPER,
    // a header, a page of code.
    inline bool RdBytes(std::uint64_t at, void* dst, std::size_t n)
    {
        __try { memcpy(dst, reinterpret_cast<const void*>(at), n); return true; }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
    inline bool RdWide(std::uint64_t at, wchar_t* dst, int chars)
    {
        __try
        {
            memcpy(dst, reinterpret_cast<const void*>(at),
                   static_cast<std::size_t>(chars) * sizeof(wchar_t));
            return true;
        }
        __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
    }
}
