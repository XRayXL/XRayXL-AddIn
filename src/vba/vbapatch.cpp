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
extern "C" void XRayVbaExitThunk(void);
extern "C" void XRayVbaEndThunk(void);

namespace vba
{
    namespace
    {
        // ---------------------------------------------------------------
        // The per-slot stub, emitted at runtime because each slot needs its own original
        // handler address and thunk pointer.
        //
        //   +00  FF 15 <rip32>     call  qword ptr [rip+thunk]
        //   +06  FF 25 <rip32>     jmp   qword ptr [rip+original]
        //   +0C  CC...             padding
        //   +10  <u64 original>
        //   +18  <u64 shared thunk>
        //
        // A real call, matched by the thunk's own `ret`: pushing the original and `ret`-ing to
        // it would forge a return address, which CET shadow stacks fast-fail on. Indirect
        // through an embedded pointer, so the stub need not sit within 2 GB of the thunk.
        // ---------------------------------------------------------------
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

        // Arm and Disarm both rewrite g_patched and the dispatch table; this makes
        // the pair mutually exclusive. The owner is kept so a contained fault, which runs no
        // destructor, can still let the gate go.
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

        // ---- rtcDoEvents -----------------------------------------------------
        // Excel runs timer macros and events from inside DoEvents, so a procedure that opens
        // while a frame waits there is not that frame's callee. A detour on the export is
        // what says so: the alternative, unwinding the stack at every entry, is the walk KB
        // hazard B warns about. Four register arguments are passed straight through, so the
        // real arity does not matter; rtcDoEvents takes no floating-point argument.
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
        int                  g_endSlots = 0;

        // ROLES ARE MATCHED WHOLE, never by first letter. "end" and "exit"
        // share one, so a first-letter test sends "end" down "exit"'s arm --
        // the DEFAULT arm -- and it is silently patched with the exit thunk,
        // reading the interpreter as though a procedure were returning.
        bool IsRole(const char* role, const char* want)
        {
            return role && std::strcmp(role, want) == 0;
        }

        // Null for an unknown role, so the arm refuses rather than patch it as an exit.
        const void* ThunkForRole(const char* role)
        {
            if (IsRole(role, "bos"))   return reinterpret_cast<const void*>(&XRayVbaBosThunk);
            if (IsRole(role, "end"))   return reinterpret_cast<const void*>(&XRayVbaEndThunk);
            if (IsRole(role, "exit"))  return reinterpret_cast<const void*>(&XRayVbaExitThunk);
            return nullptr;
        }
    }

    bool IsVbaArmed() { return g_armed; }

    // After a fault part-way through arm or disarm. Disarm restores only slots that still hold
    // our stubs, so releasing the gate makes it safe to try again; holding it would refuse every
    // arm and disarm for the life of the process.
    // XRAYXL_DIAG instrument: faults while holding the gate, as a fault part-way through
    // arm or disarm would.
    void FaultWhileArmGateHeldForProbe()
    {
        ArmGate gate;
        volatile int* nowhere = nullptr;
        *nowhere = gate.held ? 1 : 2;
    }

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

        // Nothing may still be inside a hook when the procedure table, the
        // totals and the p-code lengths are reset -- a lingering thread would
        // keep writing into slots that have just been zeroed and re-claimed.
        if (!WaitForHooksQuiet(2000))
            return "VBA tracing: REFUSED -- a previous session's hooks are still running";

        // Session state is set further down, after the last step that can refuse.

        // One stub per distinct original handler.
        std::map<std::uint64_t, std::uint8_t*> byOriginal;
        const std::size_t pageBytes = 0x1000;
        // Anywhere will do: the stub reaches the thunk through an embedded
        // absolute pointer, so there is no rel32 distance to satisfy and no scan
        // for a nearby free page.
        g_page = static_cast<std::uint8_t*>(
            VirtualAlloc(nullptr, pageBytes, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE));
        if (!g_page)
            return "VBA tracing: could not allocate the stub page";

        // THE PAGE IS RELEASED ON EVERY REFUSAL, WITHOUT ANYONE REMEMBERING TO. The
        // guard frees unless Keep() says the arm succeeded, so a refusal added later
        // cannot forget.
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
            // Unknown role: refuse the arm.
            if (shared == nullptr)
            {
                unknownRole = site.role ? site.role : "(null)";
                break;
            }
            std::uint8_t* stub = g_page + used;
            EmitStub(stub, shared, orig);   // indirect call: no distance to satisfy
            used += kStubSize;
            byOriginal[orig] = stub;
        }
        // FAIL CLOSED IF ANY WANTED SLOT HAS NO STUB. That every stub fits in one page
        // is a coincidence between two unrelated constants, so it is checked rather
        // than relied upon.
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

        // The stub page must be executable and NOT writable, and it is checked
        // rather than assumed: "impossible by current structure" is exactly the kind
        // of guarantee that quietly stops being true.
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

        // Latch the emit mode for this session: TOP emits only depth-1
        // frames, ALL emits every frame. Read here, at arm, never on the hot
        // path -- the setter refuses while armed for this reason.
        SetEmitTopLevelOnly(core::modes::GetDepth(core::modes::Source::Vba) == core::modes::Depth::Top);
        SetCapture(core::modes::GetArgs(core::modes::Source::Vba),
                   core::modes::GetRetVal(core::modes::Source::Vba));

        // OBJECTS. Nothing is derived here: QueryInterface validates itself in
        // the hook, so an interface id that is wrong costs the detail for that
        // class and can never produce a false claim. The counters say what
        // actually happened.
        ResetObjectTotals();
        SetDescribeObjects(core::modes::GetObjects(core::modes::Source::Vba));

        // From the same verified SlotSet the patching uses, because this is the
        // PRODUCER half and the consumer must never re-derive. A failure is not
        // fatal -- values still decode, types are simply absent -- but it is
        // reported rather than silently skipped.
        {
            PcodeLengths pl;
            // The lengths belong to an opcode set, not a table address, and the partition
            // fingerprint says whether this is the set kSigLength describes. Without it the
            // p-code walk is dropped rather than the arm refused: values and call timing are
            // still true.
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
                // `framedCount` 0 means the gate is OFF and attribution is
                // running unfiltered -- a different session from one where the
                // gate is on, and the two must not read alike in a log.
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

        // The VBA side may be the ONLY thing armed -- a workbook with macros and
        // no XLL add-ins is ordinary -- so it cannot rely on the XLL half having
        // opened the trace. Same file either way: one timeline, not two.
        // Open and Close are counted in csv, so this side just opens and closes.
        emit::csv::Open(core::modes::GetBufferBytes(), core::modes::GetPauseOnFull(),
                        core::modes::GetFormat());

        // Before any slot is live, so no frame opens without it.
        {
            std::string why;
            if (!HookDoEvents(why))
                core::Log::Note("VBA tracing: rtcDoEvents not hooked (" + why + "), so a macro Excel "
                                "runs inside another's DoEvents is reported beneath it");
        }

        g_bosSlots = g_exitSlots = g_endSlots = 0;
        for (const PatchSite& site : s.sites)
        {
            const std::uint64_t orig = g_base + site.handlerRva;
            std::uint64_t* slot = table + site.slot;
            if (*slot != orig) continue;              // somebody else got here first
            // find(), never operator[]: a miss must skip the slot, not invent a
            // null function pointer for it.
            const auto it = byOriginal.find(orig);
            if (it == byOriginal.end() || !it->second) continue;
            const std::uint64_t stub = reinterpret_cast<std::uint64_t>(it->second);
            g_patched.push_back({ slot, orig, stub });
            InterlockedExchange64(reinterpret_cast<volatile LONG64*>(slot), static_cast<LONG64>(stub));
            if      (IsRole(site.role, "bos"))   ++g_bosSlots;
            else if (IsRole(site.role, "end"))   ++g_endSlots;
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
        pageGuard.Keep();      // armed: the page belongs to the session now
        std::ostringstream o;
        o << "VBA tracing: ARMED [" << core::modes::DepthName(core::modes::GetDepth(core::modes::Source::Vba)) << "] -- "
          << g_patched.size() << " of " << s.sites.size() << " slots patched ("
          << g_bosSlots << " bos, " << g_exitSlots << " exit, "
          << g_endSlots << " end; "
          << byOriginal.size() << " stubs), table +0x" << std::hex << s.tableRva << std::dec;
        // SAY WHEN A FEATURE IS ABSENT. Without the `End` slot a chain killed by `End` is closed
        // only by the stack-pointer backstop, so whatever runs next can nest
        // under frames that are already dead -- the depth is inflated and the
        // parentage invented, silently.
        if (!s.endOk)
            o << " -- NO `End` HANDLING: the End slot did not verify;"
                 " depth after an End may be overstated";
        return o.str();
    }

    // From the same counters the report prints, after the session, so nothing
    // is added to the hot path.
    std::vector<std::string> UnknownOpcodeWarnings()
    {
        std::vector<std::string> w;
        if (PcodeStopWarning()[0])        w.push_back(PcodeStopWarning());
        // Kept separate from the stop warning: "no length" and "wrong length"
        // are different defects with different fixes, and the second is worse.
        if (PcodeSuspectWarning()[0])     w.push_back(PcodeSuspectWarning());
        if (ArgTypeUnknownWarning()[0])   w.push_back(ArgTypeUnknownWarning());
        const std::string r = ReturnTypeUnknownWarning();
        if (!r.empty())                   w.push_back(r);
        return w;
    }

    // THE P-CODE DIAGNOSTICS, logged at disarm from the counters the walk
    // kept. The levels are deliberate: WARNING is the shape fuzzer's oracle,
    // so only a defect goes there; coverage news is INFO; noise about noise
    // is DEBUG.
    void LogDisarmDiagnostics()
    {
        // WHAT THE TRACER MET AND DID NOT UNDERSTAND. An unknown opcode does
        // not make the trace wrong, it makes it QUIETLY INCOMPLETE: "(?)" and
        // an empty ret read as findings rather than gaps. Silence is the
        // assertable state.
        for (const std::string& w : UnknownOpcodeWarnings()) core::Log::Warning(w);

        // The table's health. A correct table walks every procedure cleanly, so a shortfall is
        // a defect. One wrong length in a common opcode drops the clean rate sharply while
        // blaming nothing, because every break follows a resync.
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

        // Lengths this session used that nothing has confirmed: coverage news,
        // not a defect -- the walk did not break on them -- and the one channel
        // through which a real workbook says which lengths to measure next.
        if (const char* uv = PcodeUnverifiedWarning())
            if (uv[0]) core::Log::Note(uv);

        // The p-code corpus: it opens a file, so here and never in a hook. It
        // is the evidence the length table is checked against.
        if (core::modes::DiagEnabled())
        {
            wchar_t leaf[64];
            _snwprintf_s(leaf, _TRUNCATE, L"\\XRayXL_pcode_%lu.txt", GetCurrentProcessId());
            core::Log::Note(WritePcodeCorpus(core::EnsureAppSubdir(L"Logs") + leaf));
        }

        // What the frame gate refused: noise on a healthy build, unless a
        // NAMED LOAD was refused, which means the frame scan itself is wrong.
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

        // The stub page is leaked: a VBA thread may still be inside a stub.
        // A thread past the slot load but before inHook++ is not seen by WaitForHooksQuiet.
        const bool quiet = WaitForHooksQuiet(2000);

        if (quiet)
        {
            FlushOpenFrames();      // no entry row is left without its exit
            // NOT ClearArmedLengths() here: the report below asks whether a length
            // table exists. Cleared after it.
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
        // The instruction stream of one procedure whose parameters did not all
        // resolve -- whether an unrecovered type is a gap in the TABLES or the
        // WALK. Developer evidence, gated behind XRAYXL_DIAG (core::modes::DiagEnabled).
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
