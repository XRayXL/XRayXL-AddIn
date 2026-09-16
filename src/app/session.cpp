#include "session.h"
#include "xll/xlltrace.h"
#include "vba/vbaderive.h"
#include "vba/vbapatch.h"
#include "core/log.h"
#include "emit/csv.h"
#include "core/render.h"

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
        core::Log::Note(vba::ArmCounting());
        return xll::Arm();
    }

    void Disarm(bool shuttingDown)
    {
        // THE OUTPUT RING, REPORTED FIRST -- the VBA disarm below tears the
        // ring down, after which capacity and drops read as nothing. The count
        // is already final: drops are producer-counted at deposit and the calc
        // has finished. Since the file carries no drop marker, this line --
        // on every path, XLL-only, VBA-only or both -- plus the Disarm return
        // value and the status summary are the ONLY places loss is reported.
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

        // VBA first and unconditionally: it can be armed on a session where the
        // XLL side armed nothing, and its p-code diagnostics come from the same
        // counters at their own log levels.
        core::Log::Note(vba::DisarmCounting());
        vba::LogDisarmDiagnostics();

        xll::Disarm(shuttingDown);
    }

    bool IsArmed() { return xll::IsArmed() || vba::IsVbaArmed(); }
}
