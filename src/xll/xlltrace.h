#pragma once
#include <windows.h>
#include <string>

// The two recorders xllthunk.asm calls, and the arming that puts them on a
// function.

namespace xll
{
    // SEH-guarded internally: a fault while decoding somebody else's memory
    // must cost the row, not the process.
    extern "C" void XRayOnEntry(void* target, void* regs);
    extern "C" void XRayOnExit (void* target, void* regs);

    struct ArmReport
    {
        int registered = 0;     // rows the registration table gave us
        int armed = 0;
        int declined = 0;
        std::string detail;     // human-readable, for the log and the dialog
    };

    // Arms the XLL source only (app/session.h arms both). Never at load, so by the time a user
    // asks every other add-in has registered.
    ArmReport Arm();
    // shuttingDown: skip the coverage check, which needs the object model.
    void      Disarm(bool shuttingDown = false);

    // Registrations that exist but are not hooked -- i.e. added after arming.
    // -1 if it could not be established.
    int       UnhookedRegistrations();
    bool      IsArmed();

    // Latched at arm from core::modes::, so the hot path never sees them change.
    void      SetCapture(bool args, bool retval);
    // DEPTH=TOP: emit only the outermost call on a thread.
    void      SetTopOnly(bool topOnly);

    // Exits dropped because the recorder was re-entered; each entry then reads like a call
    // that never returned.
    long long ExitsDropped();
    long long RecorderFaults();
    long long FramesResynced();
}
