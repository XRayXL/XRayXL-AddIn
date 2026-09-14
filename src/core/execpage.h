#pragma once
#include <windows.h>
#include <cstddef>
#include <cstdio>
#include "core/crashlog.h"
#include "core/log.h"

// Executable pages for hook stubs. Each is logged and registered with the crash
// handler, so a fault inside one is named instead of reading "<no module>".
namespace core
{
    inline void NoteExecPage(void* p, std::size_t bytes, const char* what, const char* protection)
    {
        char b[160];
        _snprintf_s(b, sizeof(b), _TRUNCATE, "%s: 0x%llX .. 0x%llX (%zu bytes, %s)", what,
                    static_cast<unsigned long long>(reinterpret_cast<ULONG_PTR>(p)),
                    static_cast<unsigned long long>(reinterpret_cast<ULONG_PTR>(p) + bytes),
                    bytes, protection);
        core::Log::Note(b);
        core::crashlog::NoteExecRegion(p, bytes, what);
    }

    // Changes a page's protection and flushes the instruction cache.
    inline bool ProtectExecPage(void* p, std::size_t bytes, DWORD protect)
    {
        DWORD old = 0;
        if (!VirtualProtect(p, bytes, protect, &old)) return false;
        FlushInstructionCache(GetCurrentProcess(), p, bytes);
        return true;
    }
}
