// The VBA disarm report. Every string here is read by a suite or the shape fuzzer's oracle,
// so changing one changes a contract.
#include "vbatrace.h"
#include "vbatrace_internal.h"
#include "vbaidentity.h"
#include "vbaobject.h"
#include "core/tracemodes.h"

#include <windows.h>
#include <algorithm>
#include <cstdio>
#include <sstream>
#include <vector>

namespace vba
{
    std::string TotalsLine()
    {
        const Totals t = ReadTotals();
        std::ostringstream o;
        o << "VBA trace: statements=" << t.statements
          << " exits=" << t.exits
          << " transitions=" << t.transitions
          << " procedures=" << t.procedures
          << " maxDepth=" << t.maxDepth
          << " recursions=" << t.recursions
          << " faults=" << t.faults
          << " stackGrowFailures=" << t.stackGrowFailures
          << " tableFull=" << t.tableFull
          << " unmatchedExits=" << t.unmatchedExits
          << " exitClosed=" << t.exitClosed
          << " exitNoMatch=" << t.exitNoMatch
          << " returnsRead=" << t.returnsRead
          << " returnsDeclined=" << t.returnsDeclined
          << " returnsOff=" << t.returnsOff
          // Four numbers, so "nothing changed" and "never looked" do not share a cell.
          << " byrefEligible=" << t.byrefEligible
          << " byrefChanged=" << t.byrefChanged
          << " byrefSame=" << t.byrefSame
          << " byrefDeclined=" << t.byrefDeclined
          << " hookFaults=" << t.hookFaults
          << " hookReentries=" << t.hookReentries
          << " stackOverflows=" << t.stackOverflows
          << " regReadFailures=" << t.regReadFailures
          << (t.tripped ? " BREAKER=OPEN(tracing stood down)" : "")
          << " framesOpened=" << t.framesOpened
          << " framesClosed=" << t.framesClosed
          << " spSame=" << t.sameTrailerSameSp
          << " spDeeper=" << t.sameTrailerDeeper
          << " spShallower=" << t.sameTrailerShallower
          << " ipEntries=" << t.ipEntries
          << " ipStaleClosed=" << t.ipStaleClosed
          << " ipUnavailable=" << t.ipUnavailable
          << " ipIntraProc=" << t.ipIntraProc
          << " ipLateOpen=" << t.ipLateOpen
          << " breakpointStops=" << t.breakpointStops
          << " stopStatements=" << t.stopStatements
          << " exitNotProcedureEnd=" << t.exitNotProcedureEnd
          << " exitOpUnreadable=" << t.exitOpUnreadable
          << " callerCell=" << t.callerCell
          << " callerOther=" << t.callerOther
          << " callerUnavailable=" << t.callerUnavailable
          << " callerFaults=" << t.callerFaults
          // threw == handled when every error was caught; a shortfall reached the top.
          << " threw=" << t.threw
          << " unwound=" << t.unwound
          << " handled=" << t.handled
          << " errEscaped=" << t.errEscaped
          << " doEventsChains=" << t.doEventsChains
          << " returnsUnmapped=" << t.returnsUnmapped
          << " closedByBackstop=" << t.closedByBackstop
          << " closedByFlush=" << t.closedByFlush
          << " closedByEnd=" << t.closedByEnd;

        // Object describer counts: objDescribed falling means an interface id stopped matching.
        {
            long long described = 0, namedOnly = 0, unknown = 0;
            ObjectTotals(described, namedOnly, unknown);
            if (described || namedOnly || unknown)
                o << " objDescribed=" << described
                  << " objNamedOnly=" << namedOnly
                  << " objUnknown=" << unknown;
        }
        // Bounded timing is not a measurement, and a reader not told so will believe the number.
        if (t.closedByBackstop || t.closedByFlush)
            o << " NOTE: " << (t.closedByBackstop + t.closedByFlush)
              << " exit row(s) carry an UPPER BOUND for ticks, not a measurement"
                 " (closed= says which)";
        // A non-cell caller is an answer; Excel declining and the call faulting are not.
        if (t.callerUnavailable)
            o << " WARNING: Excel declined to name the caller for "
              << t.callerUnavailable << " activation(s)";
        if (t.callerFaults)
            o << " WARNING: asking who called faulted " << t.callerFaults << " time(s)";
        if (t.stackGrowFailures)
            o << " WARNING: a frame stack could not grow, so VBA tracing stood down";
        // A silently degraded boundary reads exactly like one that was never wrong.
        if (t.ipUnavailable)
            o << " WARNING: the p-code activation boundary could not be evaluated"
                 " for " << t.ipUnavailable << " statement(s); those fell back to"
                 " the rsp/exit-opcode heuristic";
        // Always printed: an honest gap, and what the class-returns-every-type suite reads.
        UnmappedExitOps().Print(o, "unmappedExitOp", false);
        // Raw opcode samples and the identity walk are developer evidence, so DIAG only.
        const bool diag = core::modes::DiagEnabled();
        if (diag)
        {
            ExitOpsSeen().Print(o, "exitOpSeen", false);
            RetStores().Print(o, "retStore", true);
        }
        int named = 0, unnamed = 0;
        for (int i = 0; i < Procs().Size(); ++i)
        {
            Proc* p = Procs().At(i);
            if (p->trailer) { if (p->function[0]) ++named; else ++unnamed; }
        }
        o << " named=" << named << " unnamed=" << unnamed;
        // `named`/`unnamed` count table slots, so a procedure that never got one is invisible to
        // both. `tableFull` counts frame pushes, not procedures.
        if (t.tableFull)
            o << " WARNING: the procedure table (" << Procs().Size() << " entries) filled;"
                 " " << t.tableFull << " frame push(es) could not be recorded, so those"
                 " rows carry no module or function name and their per-procedure"
                 " counts are missing. Trace fewer procedures per arming session.";
        for (int i = 0; i < static_cast<int>(IdDecline::Count_); ++i)
            if (IdDeclineCount(static_cast<IdDecline>(i)))
                o << " id[" << IdDeclineName(static_cast<IdDecline>(i)) << "]="
                  << IdDeclineCount(static_cast<IdDecline>(i));
        if (diag && IdentityDebug() && IdentityDebug()[0]) o << "\n    " << IdentityDebug();
        return o.str();
    }

    std::string ReturnTypeUnknownWarning()
    {
        const std::uint64_t n = ReadTotals().returnsUnmapped;
        if (!n) return std::string();
        std::ostringstream o;
        o << "VBA returns: " << n << " activation(s) ended on an EXIT OPCODE with no "
             "mapping, so `ret` and `rettype` were left EMPTY rather than guessed --";
        UnmappedExitOps().Print(o, "op", false);
        o << ". Each is one case in vbaretdecode.cpp ExitReturnKind; the store-scan "
             "fallback could not name it either.";
        return o.str();
    }

    std::string Report(int maxRows)
    {
        LARGE_INTEGER f; QueryPerformanceFrequency(&f);
        const double freq = static_cast<double>(f.QuadPart);

        struct Row { std::uint64_t trailer, calls, stmts, ticks, depth;
                     const char* mod; const char* fn; };
        std::vector<Row> rows;
        for (int i = 0; i < Procs().Size(); ++i)
        {
            Proc* p = Procs().At(i);
            if (p->trailer && p->calls)       // never called: the editor's wrapper
                rows.push_back({ p->trailer, p->calls, p->statements, p->ticks,
                                 p->maxDepth, p->qualModule, p->function });
        }
        std::sort(rows.begin(), rows.end(),
                  [](const Row& a, const Row& b) { return a.ticks > b.ticks; });

        std::ostringstream o;
        o << "  " << rows.size() << " VBA procedure(s) seen\n";
        o << "      procedure                                     calls  statements       ms  depth\n";
        int n = 0;
        for (const Row& r : rows)
        {
            if (n++ >= maxRows) break;
            // An unnamed procedure says so, rather than appearing as a blank row.

            char who[640];
            if (r.fn && r.fn[0])
                std::snprintf(who, sizeof(who), "%s%s%s",
                              (r.mod && r.mod[0]) ? r.mod : "", (r.mod && r.mod[0]) ? "." : "", r.fn);
            else
                std::snprintf(who, sizeof(who), "0x%llX (unnamed)",
                              static_cast<unsigned long long>(r.trailer));

            char line[768];
            std::snprintf(line, sizeof(line),
                          "      %-44s %6llu %11llu %8.2f %6llu\n", who,
                          static_cast<unsigned long long>(r.calls),
                          static_cast<unsigned long long>(r.stmts),
                          freq > 0 ? (r.ticks * 1000.0 / freq) : 0.0,
                          static_cast<unsigned long long>(r.depth));
            o << line;
        }
        return o.str();
    }
}
