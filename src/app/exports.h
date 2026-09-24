#pragma once
#include "xlcall.h"
#include "core/contained.h"
#include "core/crashlog.h"
#include "core/log.h"
#include "emit/csv.h"
#include "session.h"

// What every worksheet-callable export shares. Internal to app/: only the files that implement
// the XRayXL_* functions include it.
namespace app
{
    // #VALUE!, handed back when a body faulted. One place that says so.
    LPXLOPER12 FaultOper();

    // Static, because Excel reads the XLOPER12 after we return; these run one at a time, on
    // whatever thread Application.Run arrived on.
    LPXLOPER12 EchoStr(const wchar_t* text);

    // The setters refuse a cell caller: a formula that reconfigures the
    // tracer on every recalc is a foot-gun.
    bool CalledFromCell();

    // Either source armed, through the same function XRayXL_IsArmed answers
    // with, so the two can never disagree.
    bool AnythingArmed();

    // Every export runs through this: the body, or on a contained fault a logged line
    // and #VALUE!. `name` null keeps the fault out of the log, for the probe that faults on
    // purpose. Pointer arguments only, so __try shares its frame with nothing that needs
    // unwinding (C2712).
    template <class Fn, class... Args>
    LPXLOPER12 GuardedOper(const char* name, Fn body, Args... args)
    {
        __try
        {
            return body(args...);
        }
        __except (core::contained::Note(GetExceptionInformation(), "XRayXL function",
                                        name ? name : "an export", "contained; #VALUE! returned"))
        {
            // No destructor ran in this unwind, so a lock the body held is still held.
            app::ReleaseHeldByThisThread();
            if (name) core::contained::Report();
            return FaultOper();
        }
    }
}
