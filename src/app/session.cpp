#include "session.h"
#include "xll/xlltrace.h"
#include "vba/vbaderive.h"
#include "vba/vbapatch.h"
#include "core/log.h"
#include "core/notify.h"
#include "core/excel_om.h"
#include "core/tracemodes.h"
#include "emit/csv.h"
#include "core/render.h"
#include "core/textvalue.h"

#include <sstream>

namespace app
{
    xll::ArmReport Arm()
    {
        if (IsArmed())
        {
            // Logged: a caller waiting on the log otherwise sees a refused arm as silence.
            core::Log::Note("arm refused: already armed -- XRayXL_Disarm first");
            xll::ArmReport r;
            r.detail = "already armed";
            return r;
        }
        // Excel loads VBA lazily and the VBA side is armed once, here: if it is wanted and absent, ask Excel for it.
        if (core::modes::VbaEnabled())
        {
            std::ostringstream vl;
            core::excelom::EnsureVbaLoaded(vl);
            if (!vl.str().empty()) core::Log::Note(vl.str());
        }

        core::Log::Note(vba::ArmCounting());
        core::ResetValuesRefused();
        xll::ArmReport r = xll::Arm();
        // Armed state is mirrored by the ribbon, which may not be what asked.
        core::NotifyStateChanged();
        return r;
    }

    void Disarm(bool shuttingDown)
    {
        // The ring is reported first: the VBA disarm below tears it down, after which capacity and
        // drops read as nothing. The file carries no drop marker, so this line, the Disarm return
        // value and the status summary are the only places loss is reported.
        if (emit::csv::RingCapacityBytes() > 0)
        {
            const long long drops  = emit::csv::RingDrops();
            const long long pauses = emit::csv::RingPauses();
            const std::size_t capB = emit::csv::RingCapacityBytes();
            char size[32];
            core::FormatBytes(capB, size, sizeof(size));
            std::ostringstream r;
            r << "output ring: " << size << " capacity, "
              << drops << " dropped, " << pauses << " hot-path pause(s)";
            if (drops > 0)
                r << " -- BUFFERWHENFULL=DROP: raise the buffer (XRayXL_SetTraceParam \"BUFFERSIZE\" <MB>)"
                     " to hold the burst, set BUFFERWHENFULL=PAUSE to lose nothing, or 0 for the synchronous write";
            else if (pauses > 0)
                r << " -- BUFFERWHENFULL=PAUSE throttled the calc; raise the buffer to reduce it";
            core::Log::Note(r.str());
        }

        // An array over the value limit keeps its shape and loses its contents; the file shows
        // it only as a header with no braces, so it is counted here too.
        if (const long long refused = core::ValuesRefused())
        {
            char size[32];
            core::FormatBytes(core::kMaxValueBytes, size, sizeof(size));
            std::ostringstream v;
            v << "values: " << refused << " array(s) over the " << size
              << " value limit written as their shape only, with no contents";
            core::Log::Note(v.str());
        }

        // VBA first and unconditionally: it can be armed on a session where the
        // XLL side armed nothing, and its p-code diagnostics come from the same
        // counters at their own log levels.
        core::Log::Note(vba::DisarmCounting());
        vba::LogDisarmDiagnostics();

        xll::Disarm(shuttingDown);

        // not while shutting down: the subscriber is the ribbon, and that would call into Office mid-teardown
        if (!shuttingDown) core::NotifyStateChanged();
    }

    bool IsArmed() { return xll::IsArmed() || vba::IsVbaArmed(); }

    void ReleaseHeldByThisThread()
    {
        core::Log::ReleaseHeldByThisThread();
        emit::csv::ReleaseHeldByThisThread();
        if (vba::ReleaseArmGateHeldByThisThread())
            core::Log::Warning("VBA tracing: an arm or disarm faulted part-way; its gate is "
                               "released so XRayXL_Disarm can be run again to restore the table");
    }
}
