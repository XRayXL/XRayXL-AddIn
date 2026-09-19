#include "contained.h"
#include "crashlog.h"
#include "log.h"

#include <cstdio>
#include <cwchar>

namespace core
{
namespace contained
{
    namespace
    {
        __declspec(thread) char t_line[512];
    }

    int Note(EXCEPTION_POINTERS* xp, const char* who, const char* where, const char* after)
    {
        const EXCEPTION_RECORD* er = (xp && xp->ExceptionRecord) ? xp->ExceptionRecord : nullptr;
        const void* addr = er ? er->ExceptionAddress : nullptr;
        char mod[MAX_PATH] = "no module (private memory)";
        HMODULE h = nullptr;
        if (addr && GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                                       GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                       reinterpret_cast<LPCWSTR>(addr), &h) && h)
        {
            wchar_t w[MAX_PATH] = {};
            GetModuleFileNameW(h, w, MAX_PATH);
            const wchar_t* leaf = wcsrchr(w, L'\\');
            leaf = leaf ? leaf + 1 : w;
            _snprintf_s(mod, _TRUNCATE, "%ls+0x%llX", leaf,
                        static_cast<unsigned long long>(
                            reinterpret_cast<const unsigned char*>(addr) -
                            reinterpret_cast<const unsigned char*>(h)));
        }
        _snprintf_s(t_line, _TRUNCATE, "%s FAULTED in %s: code 0x%08lX at %p (%s) -- %s",
                    who, where, er ? er->ExceptionCode : 0UL, addr, mod, after);
        return EXCEPTION_EXECUTE_HANDLER;
    }

    void Report()
    {
        crashlog::Note(t_line);
        Log::Error(t_line);
    }
}
}   // namespace core
