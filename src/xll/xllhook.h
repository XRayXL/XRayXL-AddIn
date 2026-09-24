#pragma once
#include "xlltypeplan.h"
#include "core/moduleid.h"
#include <windows.h>
#include <string>

// Installs the trace on a registered XLL function: an inline detour on the address Excel calls,
// whatever code is there. Nothing here knows about any add-in framework.

namespace xll
{
    struct Target
    {
        // Shared with xllthunk.asm: these two stay first, in order; the static_asserts below pin it.
        void*    original = nullptr;     // +0x00 MinHook trampoline
        INT32    argCount = 0;           // +0x08 ABI argument slots (= plan.slotCount)
        INT32    reserved = 0;           // +0x0C
        // --------------------------------------------------------------------

        // identity
        char     name[96]{};        // the name a user would recognise
        char     procName[256]{};   // the export name; Excel allows 255 characters
        char     module[64]{};

        // mechanism
        void*    exportAddr = nullptr;   // what GetProcAddress returned; what we hook
        void*    stub = nullptr;         // our per-function entry point

        // decoding
        Plan     plan;

        // stats, so a quiet trace can be told from a broken one
        volatile LONG64 calls = 0;

        // Enabled in this arming session. A slot outlives its session: the detour
        // is disabled, never removed, and the same export reuses the slot.
        bool live = false;

        // What the detour was made for, so no write lands on code that is no longer
        // it: the image, and the export's first bytes before it was patched.
        core::ModuleId image;
        unsigned char  prologue[16]{};
    };

    static_assert(offsetof(Target, original) == 0,
                  "xllthunk.asm reads TGT_ORIGINAL at [rbx+0]");
    static_assert(offsetof(Target, argCount) == 8,
                  "xllthunk.asm reads TGT_ARGCOUNT at [rbx+8]");

    // All-or-nothing per target, and never throw. Install creates the detour and queues its enable;
    // ApplyQueued patches the batch under one thread freeze rather than one per target.
    bool Install(Target* t, std::string& why);
    bool ApplyQueued(std::string& why);
    // Disables every detour; none is removed. Returns how many stayed enabled.
    int  DisableAll();
    // Frees the most recent slot after a failed Install.
    void Release(Target* t);
    // Disables what a batch hooked when its apply failed.
    int  Withdraw(Target* const* targets, int count);

    // A target live in this session whose add-in was unloaded and loaded again at
    // the same address carries no detour: patch it again, queued for the next apply.
    enum class Reloaded { No, Repatched, Refused };
    Reloaded RepatchIfReloaded(Target* t, std::string& why);

    // A fixed array, so a Target* captured by a running thunk can never dangle mid-call.
    Target* Allocate();
    void Publish(Target* t);

    // Every target we armed, for the report and for Disarm.
    int  Count();
    Target* At(int i);
    // The target hooked at this export in this session, or nullptr -- what
    // "already hooked" and the coverage check both ask.
    Target* FindTarget(void* exportAddr);
    // The slot an earlier session hooked this export in, kept for reuse.
    Target* FindRetired(void* exportAddr);

    // Each reason separate, so "never looked" can be told from "looked and
    // found nothing".
    struct Declines
    {
        long moduleNotLoaded = 0;
        long procNotFound = 0;
        long typeTextUnparsed = 0;
        long detourFailed = 0;
        long stubSpaceExhausted = 0;
        long ownModule = 0;
    };
    Declines& DeclineCounts();

    // Arming's per-function time, split because a total cannot say which part to fix.
    struct InstallCost
    {
        long long stubUs   = 0;   // MakeStub: our 24-byte entry stub
        long long createUs = 0;   // MH_CreateHook: MinHook's trampoline + disasm
        long long queueUs  = 0;   // MH_QueueEnableHook: bookkeeping only
        int       calls    = 0;
    };
    InstallCost TakeInstallCost();

    void ResetDeclines();
}
