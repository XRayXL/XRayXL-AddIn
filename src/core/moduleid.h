#pragma once
#include "core/safemem.h"
#include <windows.h>
#include <cstdint>

// The image holding an address, read without taking a reference on it. No other
// DLL is pinned (D90), so a write into one first checks its image is still there.
namespace core
{
    struct ModuleId
    {
        HMODULE base  = nullptr;
        DWORD   stamp = 0;      // link timestamp
        DWORD   size  = 0;      // SizeOfImage

        bool operator==(const ModuleId& o) const { return base == o.base && stamp == o.stamp && size == o.size; }
        bool operator!=(const ModuleId& o) const { return !(*this == o); }
    };

    inline bool ModuleAt(const void* addr, ModuleId& out)
    {
        HMODULE h = nullptr;
        if (!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                reinterpret_cast<LPCWSTR>(addr), &h) || h == nullptr)
            return false;
        const std::uint64_t base = reinterpret_cast<std::uint64_t>(h);
        IMAGE_DOS_HEADER dos{};
        if (!RdBytes(base, &dos, sizeof dos) || dos.e_magic != IMAGE_DOS_SIGNATURE) return false;
        IMAGE_NT_HEADERS64 nt{};
        if (!RdBytes(base + static_cast<std::uint32_t>(dos.e_lfanew), &nt, sizeof nt) ||
            nt.Signature != IMAGE_NT_SIGNATURE) return false;
        out.base  = h;
        out.stamp = nt.FileHeader.TimeDateStamp;
        out.size  = nt.OptionalHeader.SizeOfImage;
        return true;
    }
}
