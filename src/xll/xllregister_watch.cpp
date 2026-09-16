#include "xllregister_watch.h"
#include "core/excel_api.h"
#include "xlcall.h"
#include "MinHook.h"
#include "core/log.h"
#include "core/xltype.h"

#include <windows.h>
#include <cstring>
#include <cstdio>
#include <atomic>
#include <mutex>
#include <vector>

namespace xll
{
    namespace regwatch
    {
        namespace
        {
            typedef int(__stdcall* Excel12Proc)(int xlfn, int coper,
                                                LPXLOPER12* rgpxloper12,
                                                LPXLOPER12 xloper12Res);

            Excel12Proc g_original = nullptr;
            void*       g_target = nullptr;
            // Whether the detour is enabled. It is created once and never removed.
            std::atomic<bool> g_watching{ false };
            std::atomic<ApplyFn> g_apply{ nullptr };   // read by the worker, cleared by Remove

            // A LOCK IS FINE HERE: registrations are rare -- a burst at add-in
            // load and nothing between -- and the hot path is the pass-through
            // below, which takes none. The worker copies the batch out and
            // RELEASES the lock before patching, so it is never held across
            // MinHook's thread suspension.
            std::mutex g_mutex;
            std::vector<Captured> g_pending;

            // WHEN THE FIRST OF A BATCH WAS SEEN. Between a registration and
            // its patch the function can be CALLED and not traced -- a macro
            // that loads an add-in and calculates 4ms later misses every call,
            // 6 of 6. Closing that window means patching per registration
            // (2,829ms against 31ms), so it is narrowed and REPORTED instead.
            ULONGLONG g_firstPendingTick = 0;

            // Short enough that the untraced window is small, long enough that
            // a registration burst still batches into one patch.
            const DWORD kPollMs = 25;

            HANDLE g_stop = nullptr;
            HANDLE g_worker = nullptr;

            volatile LONG64 g_calls = 0;
            volatile LONG64 g_registers = 0;
            volatile LONG64 g_hooked = 0;
            volatile LONG64 g_faults = 0;
            volatile LONG64 g_declinedCut = 0;

            // False if truncated.
            bool CopyCounted(const XLOPER12* x, wchar_t* dst, size_t cap)
            {
                dst[0] = 0;
                if (x == nullptr) return true;
                if ((x->xltype & core::kXlTypeMask) != xltypeStr || x->val.str == nullptr) return true;
                const size_t len = static_cast<size_t>(x->val.str[0]);
                size_t n = len;
                if (n >= cap) n = cap - 1;
                memcpy(dst, x->val.str + 1, n * sizeof(wchar_t));
                dst[n] = 0;
                return n == len;
            }

            // Its own frame: __try may not share one with anything needing
            // unwinding, and this must never throw into an add-in's C API call.
            void ReadArgs(Captured& c, int coper, LPXLOPER12* rgp, LPXLOPER12 res, int rc)
            {
                __try
                {
                    // xlfRegister(module, procedure, typeText, functionText, ...). A field cut
                    // short is unusable: a cut name can spell another export, and a cut type
                    // text parses to a shorter arity.
                    const bool whole =
                        (coper <= 0 || rgp == nullptr || CopyCounted(rgp[0], c.module,       _countof(c.module)))    &&
                        (coper <= 1 || rgp == nullptr || CopyCounted(rgp[1], c.procedure,    _countof(c.procedure))) &&
                        (coper <= 2 || rgp == nullptr || CopyCounted(rgp[2], c.typeText,     _countof(c.typeText)))  &&
                        (coper <= 3 || rgp == nullptr || CopyCounted(rgp[3], c.functionText, _countof(c.functionText)));
                    if (!whole)
                    {
                        InterlockedIncrement64(&g_declinedCut);   // a correct decline, not a fault
                        c.procedure[0] = 0;      // unusable; the worker skips it
                        return;
                    }
                    c.refused = (rc != xlretSuccess);
                    if (res != nullptr)
                    {
                        const int t = res->xltype & core::kXlTypeMask;
                        if (t == xltypeNum)      c.id = res->val.num;
                        else if (t == xltypeInt) c.id = res->val.w;
                        else                     c.refused = true;   // an error where the id would be
                    }
                }
                __except (EXCEPTION_EXECUTE_HANDLER)
                {
                    InterlockedIncrement64(&g_faults);
                    c.procedure[0] = 0;          // unusable; the worker skips it
                }
            }

            int __stdcall Hook(int xlfn, int coper, LPXLOPER12* rgp, LPXLOPER12 res)
            {
                InterlockedIncrement64(&g_calls);

                // PASS THROUGH FIRST, ALWAYS: the add-in's call must behave
                // exactly as it would have, and the registration must have
                // happened before its id can be read out of the result.
                const int rc = g_original(xlfn, coper, rgp, res);

                // The xlfn carries flags in the high bits, so compare the
                // opcode alone.
                if ((xlfn & core::kXlFunctionMask) == xlfRegister)
                {
                    InterlockedIncrement64(&g_registers);
                    Captured c{};
                    ReadArgs(c, coper, rgp, res, rc);
                    if (c.procedure[0] != 0)
                    {
                        // Nothing may throw into the add-in's C API call.
                        try
                        {
                            std::lock_guard<std::mutex> lock(g_mutex);
                            if (g_pending.empty()) g_firstPendingTick = GetTickCount64();
                            g_pending.push_back(c);
                        }
                        catch (...) { InterlockedIncrement64(&g_faults); }
                    }
                }
                return rc;
            }

            // ONE EVENT, ONE TIMEOUT: waking on a timer rather than a signal
            // keeps the hook free of syscalls. The interval IS the exposure
            // window, so it is short -- a burst of 47 registrations still lands
            // in one batch, arriving within a millisecond of each other.
            DWORD WINAPI WorkerProc(LPVOID)
            {
                for (;;)
                {
                    const bool stopping =
                        (WaitForSingleObject(g_stop, kPollMs) == WAIT_OBJECT_0);

                    std::vector<Captured> batch;
                    {
                        std::lock_guard<std::mutex> lock(g_mutex);
                        batch.swap(g_pending);
                    }
                    const ApplyFn apply = g_apply.load();
                    if (!batch.empty() && apply != nullptr)
                    {
                        const ULONGLONG waited = GetTickCount64() - g_firstPendingTick;
                        InterlockedAdd64(&g_hooked,
                                         apply(&batch[0], static_cast<int>(batch.size())));
                        char b[192];
                        _snprintf_s(b, _TRUNCATE,
                                    "register watch: applied %d capture(s) %llums after the "
                                    "first was seen -- calls in that window were NOT traced",
                                    static_cast<int>(batch.size()), waited);
                        core::Log::Note(b);
                    }
                    if (stopping) break;         // drain once more, then leave
                }
                return 0;
            }
        }

        bool Install(ApplyFn apply, std::string& why)
        {
            if (g_watching) return true;                 // already watching
            g_apply = apply;
            {
                // What a disabled detour's stragglers captured belongs to no session.
                std::lock_guard<std::mutex> lock(g_mutex);
                g_pending.clear();
            }

            // CREATED ONCE AND NEVER REMOVED. A thread still inside the detour calls
            // through g_original, so its trampoline has to outlive every disarm.
            if (g_original == nullptr)
            {
                HMODULE self = GetModuleHandleW(nullptr);
                if (self == nullptr) { why = "no EXCEL.EXE module handle"; return false; }
                g_target = reinterpret_cast<void*>(GetProcAddress(self, "MdCallBack12"));
                if (g_target == nullptr) { why = "EXCEL.EXE exports no MdCallBack12"; return false; }

                // MinHook may not be initialised yet: arming initialises it only when
                // something is registered. The status is printed.
                const MH_STATUS init = MH_Initialize();
                if (init != MH_OK && init != MH_ERROR_ALREADY_INITIALIZED)
                {
                    char b[96];
                    _snprintf_s(b, _TRUNCATE, "MH_Initialize %s", MH_StatusToString(init));
                    why = b;
                    g_target = nullptr;
                    return false;
                }

                void* trampoline = nullptr;
                const MH_STATUS cs = MH_CreateHook(g_target, reinterpret_cast<void*>(&Hook),
                                                   &trampoline);
                if (cs != MH_OK)
                {
                    // The status AND the prologue: MinHook declines a target it
                    // cannot relocate, and the instructions actually there are the
                    // only thing that says why.
                    const unsigned char* b0 = static_cast<const unsigned char*>(g_target);
                    char b[176];
                    _snprintf_s(b, _TRUNCATE,
                                "MH_CreateHook on MdCallBack12 failed (%s) at %p, "
                                "prologue %02X %02X %02X %02X %02X %02X %02X %02X",
                                MH_StatusToString(cs), g_target,
                                b0[0], b0[1], b0[2], b0[3], b0[4], b0[5], b0[6], b0[7]);
                    why = b;
                    g_target = nullptr;
                    return false;
                }
                // Set before enabling, or a call already in the detour calls null.
                g_original = reinterpret_cast<Excel12Proc>(trampoline);
            }
            if (MH_EnableHook(g_target) != MH_OK)
            {
                why = "MH_EnableHook on MdCallBack12 failed";
                return false;
            }
            g_watching = true;

            g_stop = CreateEventW(nullptr, TRUE, FALSE, nullptr);
            if (g_stop != nullptr)
                g_worker = CreateThread(nullptr, 0, WorkerProc, nullptr, 0, nullptr);
            if (g_worker == nullptr)
            {
                // Log it: the caller ignores `why` when Install returns true.
                why = "watching, but the apply thread did not start -- late "
                      "registrations will NOT be hooked";
                core::Log::Note("register watch: " + why);
            }
            return true;
        }

        void Remove()
        {
            // BEFORE unhooking, so the worker cannot be inside a patch pass
            // while the detour is taken out from under it. It drains once more
            // on the way out, so a registration seen a moment ago is not
            // silently discarded. No timeout: the worker never calls Excel, and
            // returning early would let Disarm change the target table under it.
            if (g_stop != nullptr) SetEvent(g_stop);
            if (g_worker != nullptr)
            {
                WaitForSingleObject(g_worker, INFINITE);
                CloseHandle(g_worker);
                g_worker = nullptr;
            }
            if (g_stop != nullptr) { CloseHandle(g_stop); g_stop = nullptr; }
            g_apply = nullptr;

            if (!g_watching) return;
            MH_DisableHook(g_target);      // disabled, never removed: see Install
            g_watching = false;
        }

        Stats Snapshot()
        {
            Stats s;
            s.calls     = g_calls;
            s.registers = g_registers;
            s.hooked    = g_hooked;
            s.faults    = g_faults;
            s.declinedCut = g_declinedCut;
            return s;
        }
    }
}
