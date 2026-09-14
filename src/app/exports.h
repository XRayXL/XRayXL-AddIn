#pragma once
#include "xlcall.h"
#include "core/crashlog.h"
#include "core/log.h"
#include "emit/csv.h"

// WHAT EVERY WORKSHEET-CALLABLE EXPORT SHARES. Internal to app/: the three
// files that implement the XRayXL_* functions include it, nothing else does.
namespace app
{
    // #VALUE!, handed back when a body faulted. One place that says so.
    LPXLOPER12 FaultOper();

    // A string reply. Static, because Excel reads the XLOPER12 after we
    // return; these run one at a time, on whatever thread Application.Run
    // arrived on.
    LPXLOPER12 EchoStr(const wchar_t* text);

    // The setters refuse a cell caller: a formula that reconfigures the
    // tracer on every recalc is a foot-gun.
    bool CalledFromCell();

    // Either source armed, through the same function XRayXL_IsArmed answers
    // with, so the two can never disagree.
    bool AnythingArmed();

    // EVERY EXPORT RUNS THROUGH THIS: the body, or on a contained fault a note
    // and #VALUE!. Pointer arguments only, so __try shares its frame with
    // nothing that needs unwinding (C2712).
    template <class Fn, class... Args>
    LPXLOPER12 GuardedOper(const char* faultNote, Fn body, Args... args)
    {
        __try
        {
            return body(args...);
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            // No destructor ran in this unwind, so a lock the body held is still held.
            core::Log::ReleaseHeldByThisThread();
            emit::csv::ReleaseHeldByThisThread();
            if (faultNote) core::crashlog::Note(faultNote);
            return FaultOper();
        }
    }
}
