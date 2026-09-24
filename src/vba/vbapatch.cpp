#include "core/crashlog.h"
#include "vbapatch.h"
#include "vbaobject.h"
#include "vbaretdecode.h"
#include "vbaderive.h"
#include "vbatrace.h"
#include "vbaargs.h"
#include "vbapcode.h"
#include "emit/csv.h"
#include "core/log.h"
#include "core/tracemodes.h"
#include "core/excel_api.h"
#include "core/execpage.h"
#include "core/moduleid.h"
#include "MinHook.h"

#include <windows.h>
#include <cstring>
#include <cstdio>
#include <sstream>
#include <string>
#include <vector>
#include <map>

// The shared thunks, from vbathunk.asm. Entered by CALL from a per-slot stub.
extern "C" void XRayVbaBosThunk(void);
extern "C" void XRayVbaBosBpThunk(void);
extern "C" void XRayVbaExitThunk(void);
extern "C" void XRayVbaEndThunk(void);
extern "C" void XRayVbaStopThunk(void);

namespace vba
{
    namespace
    {
        // The per-slot stub, emitted at runtime because each slot has its own original handler.
        //   +00  FF 15 <rip32>     call  qword ptr [rip+thunk]
        //   +06  FF 25 <rip32>     jmp   qword ptr [rip+original]
        //   +0C  CC...             padding
        //   +10  <u64 original>
        //   +18  <u64 shared thunk>
        // A real call matched by the thunk's `ret`: push-and-ret would forge a return address,
        // which CET shadow stacks fast-fail on. Indirect, so no 2 GB reach is needed.
        constexpr std::size_t kStubSize = 0x20;
        constexpr std::size_t kOffOrig  = 0x10;
        constexpr std::size_t kOffThunk = 0x18;

        void EmitStub(std::uint8_t* p, const void* shared, std::uint64_t original)
        {
            std::size_t i = 0;
            // call qword ptr [rip + (kOffThunk - end-of-instruction)]
            p[i++] = 0xFF; p[i++] = 0x15;
            const std::int32_t c32 = static_cast<std::int32_t>(kOffThunk - (i + 4));
            std::memcpy(p + i, &c32, 4); i += 4;

            // jmp qword ptr [rip + (kOffOrig - end-of-instruction)]
            p[i++] = 0xFF; p[i++] = 0x25;
            const std::int32_t d32 = static_cast<std::int32_t>(kOffOrig - (i + 4));
            std::memcpy(p + i, &d32, 4); i += 4;

            while (i < kOffOrig) p[i++] = 0xCC;
            std::memcpy(p + kOffOrig, &original, 8);
            const std::uint64_t thunk = reinterpret_cast<std::uint64_t>(shared);
            std::memcpy(p + kOffThunk, &thunk, 8);
        }

        struct Patched
        {
            std::uint64_t* slot = nullptr;
            std::uint64_t  original = 0;
            std::uint64_t  stub = 0;
        };

        bool                 g_armed = false;

        // Arm and Disarm are mutually exclusive. The owner is kept so a contained fault, which
        // runs no destructor, can still let the gate go.
        volatile LONG g_armBusy = 0;
        volatile LONG g_armBusyOwner = 0;

        struct ArmGate
        {
            bool held;
            ArmGate() : held(InterlockedCompareExchange(&g_armBusy, 1, 0) == 0)
            {
                if (held) InterlockedExchange(&g_armBusyOwner, static_cast<LONG>(GetCurrentThreadId()));
            }
            ~ArmGate()
            {
                if (!held) return;
                InterlockedExchange(&g_armBusyOwner, 0);
                InterlockedExchange(&g_armBusy, 0);
            }
        };
        std::uint8_t*        g_page = nullptr;
        std::vector<Patched> g_patched;
        std::uint64_t        g_base = 0;

        // ---- rtcDoEvents ----
        // A procedure opening while a frame waits in DoEvents is not its callee; a detour says so
        // without a stack walk at every entry. Four register arguments pass straight through, so
        // the real arity does not matter; rtcDoEvents takes no floating-point argument.
        using DoEventsFn = INT_PTR(__stdcall*)(INT_PTR, INT_PTR, INT_PTR, INT_PTR);
        DoEventsFn g_doEventsOrig = nullptr;
        void*      g_doEventsTarget = nullptr;
        bool       g_doEventsHooked = false;

        INT_PTR __stdcall DoEventsDetour(INT_PTR a, INT_PTR b, INT_PTR c, INT_PTR d)
        {
            INT_PTR r = 0;
            NoteDoEventsEnter();
            // __finally, because End inside DoEvents unwinds and never returns here.
            __try   { r = g_doEventsOrig(a, b, c, d); }
            __finally { NoteDoEventsLeave(); }
            return r;
        }

        // Created once and kept; enabled with the patches and disabled with them.
        bool HookDoEvents(std::string& why)
        {
            if (g_doEventsHooked) return MH_EnableHook(g_doEventsTarget) == MH_OK;

            FARPROC fp = GetProcAddress(reinterpret_cast<HMODULE>(g_base), "rtcDoEvents");
            if (fp == nullptr) { why = "rtcDoEvents not exported"; return false; }
            g_doEventsTarget = reinterpret_cast<void*>(fp);

            const MH_STATUS init = MH_Initialize();
            if (init != MH_OK && init != MH_ERROR_ALREADY_INITIALIZED)
            { why = MH_StatusToString(init); return false; }

            void* tramp = nullptr;
            const MH_STATUS ch = MH_CreateHook(g_doEventsTarget, reinterpret_cast<void*>(&DoEventsDetour), &tramp);
            if (ch != MH_OK) { why = MH_StatusToString(ch); return false; }
            g_doEventsOrig = reinterpret_cast<DoEventsFn>(tramp);

            const MH_STATUS en = MH_EnableHook(g_doEventsTarget);
            if (en != MH_OK) { why = MH_StatusToString(en); return false; }
            g_doEventsHooked = true;
            return true;
        }

        void UnhookDoEvents()
        {
            if (g_doEventsHooked) MH_DisableHook(g_doEventsTarget);
        }
        core::ModuleId       g_vbe7;        // the image armed; disarm restores into no other
        int                  g_bosSlots = 0, g_exitSlots = 0;
        int                  g_endSlots = 0, g_bosBpSlots = 0, g_stopSlots = 0;

        // Matched whole: "end" and "exit" share a first letter, and exit is the default arm.
        bool IsRole(const char* role, const char* want)
        {
            return role && std::strcmp(role, want) == 0;
        }

        // Null for an unknown role, so the arm refuses rather than patch it as an exit.
        const void* ThunkForRole(const char* role)
        {
            if (IsRole(role, "bos"))   return reinterpret_cast<const void*>(&XRayVbaBosThunk);
            if (IsRole(role, "bosbp")) return reinterpret_cast<const void*>(&XRayVbaBosBpThunk);
            if (IsRole(role, "end"))   return reinterpret_cast<const void*>(&XRayVbaEndThunk);
            if (IsRole(role, "stop"))  return reinterpret_cast<const void*>(&XRayVbaStopThunk);
            if (IsRole(role, "exit"))  return reinterpret_cast<const void*>(&XRayVbaExitThunk);
            return nullptr;
        }
    }

    bool IsVbaArmed() { return g_armed; }

    // XRAYXL_DIAG instrument: faults while holding the gate, as a fault part-way through
    // arm or disarm would.
    void FaultWhileArmGateHeldForProbe()
    {
        ArmGate gate;
        volatile int* nowhere = nullptr;
        *nowhere = gate.held ? 1 : 2;
    }

    // After a fault part-way through arm or disarm. Disarm restores only slots that still hold
    // our stubs, so releasing the gate makes it safe to try again; holding it would refuse every
    // arm and disarm for the life of the process.
    bool ReleaseArmGateHeldByThisThread()
    {
        if (InterlockedCompareExchange(&g_armBusyOwner, 0, 0) != static_cast<LONG>(GetCurrentThreadId()))
            return false;
        InterlockedExchange(&g_armBusyOwner, 0);
        InterlockedExchange(&g_armBusy, 0);
        return true;
    }

    std::string ArmCounting()
    {
        ArmGate gate;
        if (!gate.held) return "VBA tracing: arm/disarm already in progress";
        if (g_armed) return "VBA tracing: already armed";
        // Derived and logged once per arm, whether or not VBA tracing is on.
        LoadedVbe7 img;
        const SlotSet s = img.Ok() ? Derive(img) : SlotSet{};
        core::Log::Note(img.Ok() ? "VBA: " + Describe(s)
                                 : std::string("VBA: VBE7 not loaded -- nothing to derive"));

        if (!core::modes::VbaEnabled())
            return "VBA tracing: OFF -- not requested "
                   "(XRayXL_SetTraceParam(\"VBA\",\"DEPTH\",\"ALL\") to turn it on) "
                   "-- dispatch table untouched";

        if (!img.Ok()) return "VBA tracing: VBE7 not loaded -- nothing to patch";
        if (!s.found || !s.verified)
            return "VBA tracing: REFUSED to patch -- " + Describe(s);

        // Recorded, not pinned: VBE7 is not ours to hold.
        if (!core::ModuleAt(reinterpret_cast<const void*>(img.Base()), g_vbe7))
            return "VBA tracing: REFUSED -- could not identify VBE7 in memory";

        g_base = img.Base();

        // A thread still inside a hook would write into slots about to be zeroed and re-claimed.
        if (!WaitForHooksQuiet(2000))
            return "VBA tracing: REFUSED -- a previous session's hooks are still running";

        // Session state is set further down, after the last step that can refuse.

        // One stub per distinct original handler.
        std::map<std::uint64_t, std::uint8_t*> byOriginal;
        const std::size_t pageBytes = 0x1000;
        // Anywhere will do: the stub reaches the thunk through an absolute pointer.
        g_page = static_cast<std::uint8_t*>(
            VirtualAlloc(nullptr, pageBytes, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE));
        if (!g_page)
            return "VBA tracing: could not allocate the stub page";

        // Frees the page unless Keep() says the arm succeeded, so a refusal added later cannot forget.
        struct PageGuard
        {
            bool keep = false;
            void Keep() { keep = true; }
            ~PageGuard()
            {
                if (!keep && g_page)
                {
                    VirtualFree(g_page, 0, MEM_RELEASE);
                    g_page = nullptr;
                }
            }
        } pageGuard;

        std::size_t used = 0;
        bool outOfRoom = false;
        const char* unknownRole = nullptr;
        for (const PatchSite& site : s.sites)
        {
            const std::uint64_t orig = g_base + site.handlerRva;
            if (byOriginal.find(orig) != byOriginal.end()) continue;
            if (used + kStubSize > pageBytes) { outOfRoom = true; break; }
            const void* shared = ThunkForRole(site.role);
            if (shared == nullptr)
            {
                unknownRole = site.role ? site.role : "(null)";
                break;
            }
            std::uint8_t* stub = g_page + used;
            EmitStub(stub, shared, orig);
            used += kStubSize;
            byOriginal[orig] = stub;
        }
        // Fail closed if a wanted slot has no stub: one page fitting them all is a coincidence.
        if (outOfRoom)
        {
            return "VBA tracing: REFUSED -- stub page too small for every slot";
        }
        if (unknownRole)
        {
            return std::string("VBA tracing: REFUSED -- the derivation asked for a slot "
                               "with role '") + unknownRole + "', which this build has no "
                               "thunk for";
        }

        if (!core::ProtectExecPage(g_page, pageBytes, PAGE_EXECUTE_READ))
            return "VBA tracing: could not make the stub page executable";
        core::NoteExecPage(g_page, pageBytes, "VBA stub page", "PAGE_EXECUTE_READ");

        // Checked rather than assumed: the stub page must be executable and not writable.
        {
            MEMORY_BASIC_INFORMATION mbi{};
            if (!VirtualQuery(g_page, &mbi, sizeof(mbi)) ||
                mbi.Protect != PAGE_EXECUTE_READ)
            {
                return "VBA tracing: REFUSED -- stub page is not execute-read-only";
            }
        }

        std::uint64_t* table = reinterpret_cast<std::uint64_t*>(g_base + s.tableRva);
        const std::size_t tableBytes = static_cast<std::size_t>(s.slots) * 8;
        DWORD prev = 0;
        if (!VirtualProtect(table, tableBytes, PAGE_READWRITE, &prev))
            return "VBA tracing: could not make the dispatch table writable";

        // Past the last refusal. No hook can fire until the loop below writes a slot.
        SetPcodeDiagnostics(core::modes::DiagEnabled());
        SetArgTypeOpcodeDiagnostics(core::modes::DiagEnabled());
        ResetTracing();

        // Latched here, never read on the hot path; the setter refuses while armed.
        SetEmitTopLevelOnly(core::modes::GetDepth(core::modes::Source::Vba) == core::modes::Depth::Top);
        SetCapture(core::modes::GetArgs(core::modes::Source::Vba),
                   core::modes::GetRetVal(core::modes::Source::Vba));

        // Nothing to derive for objects: QueryInterface validates itself, so a wrong interface id
        // costs that class's detail and never makes a false claim.
        ResetObjectTotals();
        SetDescribeObjects(core::modes::GetObjects(core::modes::Source::Vba));

        // From the same verified SlotSet the patching uses, so the consumer never re-derives.
        // A failure costs types only, and is reported.
        {
            PcodeLengths pl;
            // The lengths belong to an opcode set, which the partition fingerprint identifies.
            // On a mismatch the walk is dropped, not the arm: values and timing are still true.
            if (!s.partitionOk)
            {
                ClearArmedLengths();
                char m[256];
                _snprintf_s(m, sizeof m, _TRUNCATE,
                    "VBA p-code: UNKNOWN OPCODE SET (partition 0x%016llX, expected the "
                    "measured one) -- argument and return TYPES are OFF for this session. "
                    "Values, names, cells and timing are unaffected. The pinned instruction "
                    "lengths were measured against a different VBE7 opcode set and would "
                    "decode this one into fiction.",
                    static_cast<unsigned long long>(s.partitionHash));
                core::Log::Warning(m);
            }
            else if (PinPcodeLengths(img, s, pl))
            {
                SetArmedLengths(pl);
                // `framedCount` 0 means attribution runs ungated, which must not read like gated.
                char m[288];
                _snprintf_s(m, sizeof m, _TRUNCATE,
                    "VBA p-code: %u length(s) pinned (%u never exercised by real "
                    "compiler output), %u slot(s) are the invalid handler, "
                    "%u of %u address the frame in %u us (parameter attribution %s)",
                    pl.pinned, pl.unverifiedCount, pl.invalidCount, pl.framedCount,
                    pl.slots, pl.scanMicros,
                    pl.framedCount ? "gated on R14" : "UNGATED -- scan found nothing");
                core::Log::Note(m);
            }
            else                                ClearArmedLengths();
        }

        // VBA may be the only thing armed, so it opens the trace itself; csv counts opens, so
        // both sides share one file and one timeline.
        emit::csv::Open(core::modes::GetBufferBytes(), core::modes::GetPauseOnFull(),
                        core::modes::GetFormat(), core::modes::BreaksColumn());

        // Before any slot is live, so no frame opens without it.
        {
            std::string why;
            if (!HookDoEvents(why))
                core::Log::Note("VBA tracing: rtcDoEvents not hooked (" + why + "), so a macro Excel "
                                "runs inside another's DoEvents is reported beneath it");
        }

        g_bosSlots = g_exitSlots = g_endSlots = g_bosBpSlots = g_stopSlots = 0;
        for (const PatchSite& site : s.sites)
        {
            const std::uint64_t orig = g_base + site.handlerRva;
            std::uint64_t* slot = table + site.slot;
            if (*slot != orig) continue;              // somebody else got here first
            // find(), never operator[]: a miss must skip the slot, not invent a null pointer.
            const auto it = byOriginal.find(orig);
            if (it == byOriginal.end() || !it->second) continue;
            const std::uint64_t stub = reinterpret_cast<std::uint64_t>(it->second);
            g_patched.push_back({ slot, orig, stub });
            InterlockedExchange64(reinterpret_cast<volatile LONG64*>(slot), static_cast<LONG64>(stub));
            if      (IsRole(site.role, "bos"))   ++g_bosSlots;
            else if (IsRole(site.role, "bosbp")) ++g_bosBpSlots;
            else if (IsRole(site.role, "end"))   ++g_endSlots;
            else if (IsRole(site.role, "stop"))  ++g_stopSlots;
            else                                 ++g_exitSlots;
        }

        // Warn if the table cannot be restored or still reads back writable.
        DWORD ignored = 0;
        if (!VirtualProtect(table, tableBytes, prev, &ignored))
        {
            char m[176];
            _snprintf_s(m, sizeof m, _TRUNCATE,
                "VBA tracing: WARNING -- the dispatch table could not be returned to "
                "protection 0x%lX (error %lu); it stays WRITABLE for this process",
                static_cast<unsigned long>(prev),
                static_cast<unsigned long>(GetLastError()));
            core::Log::Warning(m);
        }
        else
        {
            MEMORY_BASIC_INFORMATION mbi{};
            if (VirtualQuery(table, &mbi, sizeof(mbi)) && (mbi.Protect & PAGE_READWRITE))
            {
                char m[160];
                _snprintf_s(m, sizeof m, _TRUNCATE,
                    "VBA tracing: WARNING -- the dispatch table reads back as writable "
                    "(0x%lX) after restore", static_cast<unsigned long>(mbi.Protect));
                core::Log::Warning(m);
            }
        }

        g_armed = true;
        pageGuard.Keep();
        std::ostringstream o;
        o << "VBA tracing: ARMED [" << core::modes::DepthName(core::modes::GetDepth(core::modes::Source::Vba)) << "] -- "
          << g_patched.size() << " of " << s.sites.size() << " slots patched ("
          << g_bosSlots << " bos, " << g_bosBpSlots << " breakpoint, " << g_exitSlots << " exit, "
          << g_endSlots << " end, " << g_stopSlots << " stop; "
          << byOriginal.size() << " stubs), table +0x" << std::hex << s.tableRva << std::dec;
        // Say when a feature is absent: without the `End` slot, what runs after an `End` can
        // nest under dead frames, inflating depth and inventing parentage.
        if (!s.endOk)
            o << " -- NO `End` HANDLING: the End slot did not verify;"
                 " depth after an End may be overstated";
        if (!s.bosBpOk)
            o << " -- NO BREAKPOINT HANDLING: the BosBp pair did not verify;"
                 " a call whose first statements have breakpoints opens late";
        // Narrower: nothing traces differently, but a long `ticks` loses its explanation.
        if (!s.stopOk)
            o << " -- NO `Stop` COUNT: the Stop slot did not verify;"
                 " time paused at a `Stop` is in ticks but not in breaks";
        return o.str();
    }

    // After the session, from the report's counters, so nothing is added to the hot path.
    std::vector<std::string> UnknownOpcodeWarnings()
    {
        std::vector<std::string> w;
        if (PcodeStopWarning()[0])        w.push_back(PcodeStopWarning());
        // Apart from the stop warning: "no length" and "wrong length" have different fixes.
        if (PcodeSuspectWarning()[0])     w.push_back(PcodeSuspectWarning());
        if (PcodeClosureWarning()[0])     w.push_back(PcodeClosureWarning());
        if (ArgTypeUnknownWarning()[0])   w.push_back(ArgTypeUnknownWarning());
        const std::string r = ReturnTypeUnknownWarning();
        if (!r.empty())                   w.push_back(r);
        return w;
    }

    // WARNING is the shape fuzzer's oracle, so only a defect goes there; coverage news is INFO,
    // noise is DEBUG.
    void LogDisarmDiagnostics()
    {
        // An unknown opcode makes the trace quietly incomplete: "(?)" and an empty ret read as
        // findings rather than gaps.
        for (const std::string& w : UnknownOpcodeWarnings()) core::Log::Warning(w);

        // A correct table walks every procedure cleanly. One wrong length drops the clean rate
        // sharply while blaming nothing, because every break follows a resync.
        long long walks = 0, clean = 0;
        PcodeHealth(walks, clean);
        // Always said, good news included, so the number can be trusted when it is good.
        if (walks > 0)
        {
            char m[200];
            _snprintf_s(m, sizeof m, _TRUNCATE,
                "VBA p-code: %lld of %lld procedure(s) walked cleanly (offset 0 to a clean end, "
                "no resynchronisation)", clean, walks);
            core::Log::Note(m);
        }
        if (walks >= 20 && clean * 10 < walks * 9)
        {
            char m[320];
            _snprintf_s(m, sizeof m, _TRUNCATE,
                "VBA p-code: THE LENGTH TABLE DOES NOT FIT THIS BUILD -- only %lld of "
                "%lld procedure(s) walked cleanly, and a correct table walks every "
                "one. Argument and return TYPES are unreliable wherever a walk "
                "resynchronised; values, names, cells and timing are not affected. "
                "Either a pinned length is wrong or an opcode this build emits has "
                "none. Send this log and the XRAYXL_DIAG p-code corpus.", clean, walks);
            core::Log::Error(m);
        }

        // Unconfirmed lengths are coverage news, not a defect: the walk did not break on them.
        if (const char* uv = PcodeUnverifiedWarning())
            if (uv[0]) core::Log::Note(uv);

        // It opens a file, so here and never in a hook.
        if (core::modes::DiagEnabled())
        {
            wchar_t leaf[64];
            _snwprintf_s(leaf, _TRUNCATE, L"\\XRayXL_pcode_%lu.txt", GetCurrentProcessId());
            core::Log::Note(WritePcodeCorpus(core::EnsureAppSubdir(L"Logs") + leaf));
        }

        // Noise on a healthy build, unless a named load was refused: then the frame scan is wrong.
        if (const char* nf = PcodeNotFramedWarning())
        {
            if (nf[0])
            {
                if (strstr(nf, "SCAN SUSPECT")) core::Log::Warning(nf);
                else                            core::Log::Debug(nf);
            }
        }
    }

    std::string DisarmCounting()
    {
        ArmGate gate;
        if (!gate.held) return "VBA tracing: arm/disarm already in progress";
        if (!g_armed) return "VBA tracing: was not armed";

        // Unloaded while armed, VBE7 took its table and every call into our stubs with it.
        core::ModuleId now;
        const bool vbe7Here = core::ModuleAt(reinterpret_cast<const void*>(g_base), now) && now == g_vbe7;
        if (!vbe7Here && !g_patched.empty())
            core::Log::Note("VBA tracing: VBE7 was unloaded while armed -- nothing to restore");

        if (vbe7Here && !g_patched.empty())
        {
            std::uint64_t* lo = g_patched.front().slot, *hi = lo;
            for (const Patched& p : g_patched)
            {
                if (p.slot < lo) lo = p.slot;
                if (p.slot > hi) hi = p.slot;
            }
            const std::size_t bytes = static_cast<std::size_t>((hi - lo) + 1) * 8;
            DWORD prev = 0;
            if (VirtualProtect(lo, bytes, PAGE_READWRITE, &prev))
            {
                // Only a slot still holding our stub: anything else was not written by us.
                for (const Patched& p : g_patched)
                    InterlockedCompareExchange64(reinterpret_cast<volatile LONG64*>(p.slot),
                                                 static_cast<LONG64>(p.original), static_cast<LONG64>(p.stub));
                DWORD ignored = 0;
                VirtualProtect(lo, bytes, prev, &ignored);
            }
            else
            {
                // Could not unpatch: the hooks are still live, so stay armed and keep the list.
                const DWORD err = GetLastError();
                char m[224];
                _snprintf_s(m, sizeof m, _TRUNCATE,
                    "VBA tracing: DISARM FAILED -- could not make the dispatch table "
                    "writable to restore %zu slot(s) (VirtualProtect error %lu). The "
                    "hooks are STILL LIVE and the session stays armed; nothing has "
                    "been reset. Try disarming again, or close Excel.",
                    g_patched.size(), static_cast<unsigned long>(err));
                core::Log::Warning(m);
                return m;
            }
        }

        // The stub page is leaked: a thread past the slot load but before inHook++ is not seen
        // by WaitForHooksQuiet, and may still be inside a stub.
        const bool quiet = WaitForHooksQuiet(2000);

        if (quiet)
        {
            FlushOpenFrames();
            // Not ClearArmedLengths() yet: the report below asks whether a length table exists.
        }
        UnhookDoEvents();
        g_patched.clear();
        g_armed = false;
        // Skip the close if the hooks did not drain: a stub may still write a row.
        if (quiet) emit::csv::Close();

        const Totals t = ReadTotals();
        std::ostringstream o;
        if (!quiet)
            o << "VBA tracing: WARNING -- hooks did not drain in 2s; state left "
                 "intact rather than reset under a live thread\n";
        o << "VBA tracing: disarmed -- " << t.statements << " statements, "
          << t.exits << " exits observed\n"
          << "    " << TotalsLine() << "\n"
          << "    " << ArgsLine()
          << "\n"
          // Every p-code decline, including `Desynced`, which means a recovered type
          // may have been attributed to the wrong argument.
          << "    " << PcodeLine() << "\n"
          << "    " << PcodeStopLine() << "\n";
        // One untyped procedure's instruction stream, telling a table gap from a walk gap.

        if (core::modes::DiagEnabled())
        {
            if (PcodeUntypedDump()[0]) o << "    " << PcodeUntypedDump() << "\n";
        }
        else
            o << "    diagnostics: off -- set XRAYXL_DIAG=1 before starting Excel for "
                 "opcode / identity / p-code evidence in this report\n";
        o << Report();
        if (quiet) ClearArmedLengths();
        return o.str();
    }

}
