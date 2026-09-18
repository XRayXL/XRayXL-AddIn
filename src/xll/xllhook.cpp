#include "xllhook.h"
#include "xllregs.h"
#include "xlldecode.h"
#include "xlltrace.h"
#include "core/crashlog.h"
#include "core/log.h"
#include "core/clock.h"
#include "core/execpage.h"
#include "MinHook.h"
#include "core/moduleid.h"

#include <cstdio>
#include <cstring>

extern "C" void XRayXllThunk();

namespace xll
{
    namespace
    {
        const int kMaxTargets = 4096;

        Target   g_targets[kMaxTargets];
        int      g_count = 0;
        Declines g_declines;

        // Each hooked function has its own entry stub, which is what tells the shared thunk
        // which function it stands in. An absolute jump, so there is no rel32 range to worry
        // about:
        //
        //    49 BA <imm64>            mov r10, Target*
        //    FF 25 00 00 00 00        jmp qword ptr [rip+0]
        //    <imm64>                  XRayXllThunk
        const int kStubSize = 24;
        // As many as fit in one 4 KB page.
        const int kStubsPerPage = 170;
        unsigned char* g_stubs = nullptr;   // the page currently being filled
        int            g_stubsUsed = 0;     // stubs used in THAT page
        bool           g_minhookReady = false;   // never uninitialised: the module is pinned for the process

        InstallCost g_installCost;
        using core::QpcMicros;

        // WHAT SITS AT A TARGET'S ADDRESS NOW. No reference is held on another
        // add-in, so it may have unloaded, loaded again, or given its address to
        // another module. An unload between this look and the write leaves the
        // memory unmapped, and MinHook's VirtualProtect then refuses to write.
        enum class Code { Gone, Foreign, Unpatched, Patched };
        Code Inspect(const Target& t)
        {
            core::ModuleId now;
            if (!core::ModuleAt(t.exportAddr, now)) return Code::Gone;
            if (now != t.image) return Code::Foreign;
            unsigned char bytes[sizeof t.prologue];
            if (!core::RdBytes(reinterpret_cast<std::uint64_t>(t.exportAddr), bytes, sizeof bytes)) return Code::Gone;
            return std::memcmp(bytes, t.prologue, sizeof bytes) == 0 ? Code::Unpatched : Code::Patched;
        }

        void* MakeStub(Target* t)
        {
            // A stub address is never reused and stub memory never freed: a thread
            // inside a stub is invisible. Pages are execute-read except while a stub
            // is written; adding write access never removes execute from a running stub.
            const std::size_t pageBytes = static_cast<std::size_t>(kStubsPerPage) * kStubSize;
            if (g_stubs == nullptr || g_stubsUsed >= kStubsPerPage)
            {
                g_stubs = static_cast<unsigned char*>(VirtualAlloc(
                    nullptr, pageBytes, MEM_COMMIT | MEM_RESERVE, PAGE_EXECUTE_READWRITE));
                if (g_stubs == nullptr) return nullptr;
                g_stubsUsed = 0;
                core::NoteExecPage(g_stubs, pageBytes, "XLL stub page", "PAGE_EXECUTE_READ");
            }
            else if (!core::ProtectExecPage(g_stubs, pageBytes, PAGE_EXECUTE_READWRITE))
            {
                return nullptr;
            }
            unsigned char* p = g_stubs + (g_stubsUsed++ * kStubSize);
            p[0] = 0x49; p[1] = 0xBA;                       // mov r10, imm64
            *reinterpret_cast<ULONG64*>(p + 2) = reinterpret_cast<ULONG64>(t);
            p[10] = 0xFF; p[11] = 0x25;                     // jmp [rip+0]
            *reinterpret_cast<UINT32*>(p + 12) = 0;
            *reinterpret_cast<ULONG64*>(p + 16) =
                reinterpret_cast<ULONG64>(&XRayXllThunk);
            core::ProtectExecPage(g_stubs, pageBytes, PAGE_EXECUTE_READ);
            return p;
        }
    }

    Declines& DeclineCounts() { return g_declines; }
    void ResetDeclines() { g_declines = Declines(); }
    int Count() { return g_count; }
    Target* At(int i) { return (i >= 0 && i < g_count) ? &g_targets[i] : nullptr; }
    Target* FindTarget(void* exportAddr)
    {
        for (int i = 0; i < g_count; i++)
            if (g_targets[i].exportAddr == exportAddr && g_targets[i].live) return &g_targets[i];
        return nullptr;
    }
    Target* FindRetired(void* exportAddr)
    {
        for (int i = 0; i < g_count; i++)
            if (g_targets[i].exportAddr == exportAddr && !g_targets[i].live) return &g_targets[i];
        return nullptr;
    }

    bool Install(Target* t, std::string& why)
    {
        if (t == nullptr || t->exportAddr == nullptr) { why = "no address"; return false; }

        // A slot from an earlier session keeps its stub and its trampoline, never
        // freed, so enabling the detour again is all that is left to do.
        if (t->original != nullptr)
        {
            const Code c = Inspect(*t);
            if (c == Code::Gone || c == Code::Foreign)
            {
                g_declines.detourFailed++;
                why = "another module now sits where this export was hooked";
                return false;
            }
            // Unloaded while armed, disarm left MinHook recording it enabled; the
            // bytes it restores are the ones already there, so this only resyncs.
            if (c == Code::Unpatched) MH_DisableHook(t->exportAddr);
            const MH_STATUS q = MH_QueueEnableHook(t->exportAddr);
            if (q != MH_OK)
            {
                g_declines.detourFailed++;
                char buf[64]; _snprintf_s(buf, _TRUNCATE, "MH_QueueEnableHook %s", MH_StatusToString(q));
                why = buf;
                return false;
            }
            return true;
        }

        // Recorded before the detour exists, so later writes can check they still
        // land on this add-in.
        if (!core::ModuleAt(t->exportAddr, t->image) ||
            !core::RdBytes(reinterpret_cast<std::uint64_t>(t->exportAddr), t->prologue, sizeof t->prologue))
        {
            g_declines.detourFailed++;
            why = "could not read the module";
            return false;
        }

        g_installCost.calls++;
        const long long tStub = QpcMicros();
        t->stub = MakeStub(t);
        g_installCost.stubUs += QpcMicros() - tStub;
        if (t->stub == nullptr)
        {
            g_declines.stubSpaceExhausted++;
            why = "no stub space";
            return false;
        }

        // A prologue detour on the address Excel calls, never a patched call site: EXCEL.EXE
        // runs Control Flow Guard, which would reject a patched call site and terminate the
        // process. The indirect call still targets the registered address.
        if (!g_minhookReady)
        {
            // The register watch may have initialised MinHook already.
            const MH_STATUS init = MH_Initialize();
            if (init != MH_OK && init != MH_ERROR_ALREADY_INITIALIZED)
            {
                char buf[64];
                _snprintf_s(buf, _TRUNCATE, "MH_Initialize %s",
                            MH_StatusToString(init));
                why = buf;
                return false;
            }

            // MinHook suspends every thread before patching, and by default finds them with a
            // snapshot of every thread on the machine. The process-scoped enumeration is a
            // local change to the vendored copy (third_party/minhook/FORK.md); the log says
            // which one is in force.
            const MH_STATUS fm = MH_SetThreadFreezeMethod(MH_FREEZE_METHOD_FAST_UNDOCUMENTED);
            core::Log::Note(fm == MH_OK
                ? "minhook: process-scoped thread freeze (NtGetNextThread)"
                : "minhook: NtGetNextThread unavailable -- falling back to the "
                  "system-wide thread snapshot; arming will be slower");

            g_minhookReady = true;
        }
        void* tramp = nullptr;
        const long long tCreate = QpcMicros();
        const MH_STATUS s = MH_CreateHook(t->exportAddr, t->stub, &tramp);
        g_installCost.createUs += QpcMicros() - tCreate;
        if (s != MH_OK)
        {
            g_declines.detourFailed++;
            char buf[64]; _snprintf_s(buf, _TRUNCATE, "MH_CreateHook %s", MH_StatusToString(s));
            why = buf;
            return false;
        }
        t->original = tramp;

        // QUEUED, not enabled: the patch happens in ApplyQueued under one
        // thread freeze for the whole batch. Queueing still validates the
        // target, so a bad one is declined and named here.
        const long long tQueue = QpcMicros();
        const MH_STATUS q = MH_QueueEnableHook(t->exportAddr);
        g_installCost.queueUs += QpcMicros() - tQueue;
        if (q != MH_OK)
        {
            MH_RemoveHook(t->exportAddr);
            g_declines.detourFailed++;
            char buf[64]; _snprintf_s(buf, _TRUNCATE, "MH_QueueEnableHook %s", MH_StatusToString(q));
            why = buf;
            return false;
        }
        return true;
    }

    InstallCost TakeInstallCost()
    {
        const InstallCost c = g_installCost;
        g_installCost = InstallCost{};
        return c;
    }

    bool ApplyQueued(std::string& why)
    {
        if (!g_minhookReady) return true;        // nothing was ever queued
        const MH_STATUS s = MH_ApplyQueued();
        if (s != MH_OK)
        {
            char buf[64]; _snprintf_s(buf, _TRUNCATE, "MH_ApplyQueued %s", MH_StatusToString(s));
            why = buf;
            return false;
        }
        return true;
    }

    int DisableAll()
    {
        // Every slot still in its own add-in is queued, which also cancels an enable
        // an arm queued and never applied. One unloaded, or with another module at its
        // address, is left as MinHook recorded it: there is nothing there of ours to
        // restore, and a queued write that failed would stop the rest of the batch.
        int queued = 0, untouched = 0;
        for (int i = 0; i < g_count; i++)
        {
            Target& t = g_targets[i];
            if (t.exportAddr == nullptr) continue;
            const Code c = Inspect(t);
            if (c == Code::Gone || c == Code::Foreign)
            {
                if (t.live) { t.live = false; untouched++; }
                continue;
            }
            MH_QueueDisableHook(t.exportAddr);
            queued++;
        }

        // Disabled, never removed: MH_RemoveHook frees the trampoline a thread still inside the
        // detour is about to call. The slot, its stub and its trampoline are kept, and arming
        // the export again reuses them.
        const bool applied = (queued == 0) || (MH_ApplyQueued() == MH_OK);
        if (untouched > 0)
        {
            char b[128];
            _snprintf_s(b, sizeof(b), _TRUNCATE,
                        "disarm: %d detour(s) left untouched -- their add-in was unloaded while armed",
                        untouched);
            core::Log::Note(b);
        }
        int stuck = 0;
        for (int i = 0; i < g_count; i++)
        {
            Target& t = g_targets[i];
            if (!t.live) continue;
            if (!applied)
            {
                // One at a time, so one bad target cannot leave the rest enabled.
                const MH_STATUS r = MH_DisableHook(t.exportAddr);
                if (r != MH_OK && r != MH_ERROR_DISABLED)
                {
                    char b[192];
                    _snprintf_s(b, sizeof(b), _TRUNCATE, "disarm: %s!%s still hooked (MH_DisableHook %s)",
                                t.module, t.name, MH_StatusToString(r));
                    core::Log::Note(b);
                    stuck++;
                    continue;
                }
            }
            t.live = false;
        }
        if (stuck > 0)
        {
            char b[128];
            _snprintf_s(b, sizeof(b), _TRUNCATE,
                        "disarm: %d hook(s) could not be disabled and still trace", stuck);
            core::Log::Warning(b);
        }
        return stuck;
    }

    // A fixed array, so the thunk's Target* can never dangle mid-call. Slots are not recycled,
    // because a disabled detour's stub still names its slot; FindRetired gives an export its
    // own slot back.
    Target* Allocate()
    {
        if (g_count >= kMaxTargets) return nullptr;
        Target* t = &g_targets[g_count];
        *t = Target{};
        return t;       // not visible to readers until Publish
    }

    // Counts a filled slot, so a reader on another thread never sees a half-built Target.
    void Publish(Target* t)
    {
        if (t == nullptr) return;
        t->live = true;
        if (g_count < kMaxTargets && t == &g_targets[g_count]) ++g_count;
    }

    // Returns an unpublished slot; Install fails in the call that allocated it.
    void Release(Target* t)
    {
        if (t == nullptr || g_count >= kMaxTargets || t != &g_targets[g_count]) return;
        *t = Target{};
    }

    int Withdraw(Target* const* targets, int count)
    {
        // Queued disables cancel the enables the batch queued, and the apply undoes
        // any it switched on before failing. Nothing is removed (see DisableAll).
        for (int i = 0; i < count; i++)
            if (targets[i] != nullptr && targets[i]->exportAddr != nullptr)
            {
                const Code c = Inspect(*targets[i]);
                if (c != Code::Gone && c != Code::Foreign) MH_QueueDisableHook(targets[i]->exportAddr);
            }
        const bool applied = (MH_ApplyQueued() == MH_OK);
        int withdrawn = 0;
        for (int i = 0; i < count; i++)
        {
            Target* t = targets[i];
            if (t == nullptr) continue;
            if (!applied)
            {
                const MH_STATUS r = MH_DisableHook(t->exportAddr);
                if (r != MH_OK && r != MH_ERROR_DISABLED) continue;
            }
            t->live = false;
            withdrawn++;
        }
        return withdrawn;
    }

    Reloaded RepatchIfReloaded(Target* t, std::string& why)
    {
        const Code c = Inspect(*t);
        if (c == Code::Patched) return Reloaded::No;
        if (c != Code::Unpatched)
        {
            g_declines.detourFailed++;
            why = "another module now sits where this export was hooked";
            return Reloaded::Refused;
        }
        // MinHook still records it enabled; the bytes it restores are the ones there now.
        MH_DisableHook(t->exportAddr);
        const MH_STATUS q = MH_QueueEnableHook(t->exportAddr);
        if (q != MH_OK)
        {
            g_declines.detourFailed++;
            char buf[64]; _snprintf_s(buf, _TRUNCATE, "MH_QueueEnableHook %s", MH_StatusToString(q));
            why = buf;
            return Reloaded::Refused;
        }
        return Reloaded::Repatched;
    }
}
