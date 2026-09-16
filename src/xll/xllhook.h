#pragma once
#include "xlltypeplan.h"
#include "core/moduleid.h"
#include <windows.h>
#include <string>

// Installing the trace on a registered XLL function.
//
// ONE MECHANISM, FOR EVERY XLL: Excel resolves an export and calls the address,
// and cannot tell what shape the code there is. Neither do we -- an inline
// detour on that address, and MinHook deals with whatever instructions it finds.
// Nothing here knows about any add-in framework.

namespace xll
{
    struct Target
    {
        // ---- SHARED WITH xllthunk.asm. THESE TWO MUST STAY FIRST, IN ORDER ----
        //
        // The thunk reads [rbx+0] for the function to call and [rbx+8] for how
        // many stack arguments to copy. Anything placed before them is what the
        // thunk calls instead -- with a char array first, it calls the name
        // string. The static_asserts below keep that honest.
        void*    original = nullptr;     // +0x00 slot's old value, or MinHook trampoline
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

    // All-or-nothing per target, and never throw. Install CREATES the detour
    // and QUEUES its enable; ApplyQueued patches everything queued under ONE
    // thread freeze, where MH_EnableHook per target would freeze and thaw every
    // thread in the process each time.
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

    // A fixed array, so a Target* captured by a running thunk can never dangle
    // mid-call.
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

    // WHERE ARMING'S PER-FUNCTION TIME GOES. The batch shares ONE thread
    // freeze, so the rest is per-function work and a total cannot say which
    // part. At a few thousand registrations that difference decides whether this
    // design scales, so it is split rather than argued about.
    struct InstallCost
    {
        long long stubUs   = 0;   // MakeStub: the 24-byte trampoline of our own
        long long createUs = 0;   // MH_CreateHook: MinHook's trampoline + disasm
        long long queueUs  = 0;   // MH_QueueEnableHook: bookkeeping only
        int       calls    = 0;
    };
    InstallCost TakeInstallCost();

    void ResetDeclines();
}
