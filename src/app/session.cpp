#include "session.h"
#include "appevents.h"
#include "settings.h"
#include "xll/xlltrace.h"
#include "vba/vbaderive.h"
#include "vba/vbapatch.h"
#include "vba/vbatrace.h"
#include "core/clock.h"
#include "core/eventlist.h"
#include "core/log.h"
#include "core/notify.h"
#include "core/excel_om.h"
#include "core/tracemodes.h"
#include "core/textbuf.h"
#include "core/valueformat.h"
#include "emit/csv.h"
#include "core/render.h"
#include "core/textvalue.h"

#include <cstdio>
#include <sstream>
#include <string>

namespace app
{
    namespace
    {
        // The session holds the trace open itself, so its arm row is the first and its disarm row
        // the last, whichever sources arm.
        bool g_sessionOpen = false;
        // Which sources ran this session: the others' totals are an earlier session's.
        bool g_vbaRan = false, g_eventsRan = false;

        void Named(core::ValueWriter& w, const char* name, long long v)
        {
            char t[32]; _snprintf_s(t, _TRUNCATE, "%lld", v);
            w.BeginNamedArg(name, "LongLong"); w.Number("LongLong", t); w.EndArg();
        }
        void Named(core::ValueWriter& w, const char* name, bool v)
        {
            w.BeginNamedArg(name, "Boolean"); w.Bool(v); w.EndArg();
        }
        void Named(core::ValueWriter& w, const char* name, const std::string& v)
        {
            std::wstring wide(v.begin(), v.end());      // ASCII: setting names and values
            w.BeginNamedArg(name, "String");
            w.String(wide.c_str(), static_cast<int>(wide.size()), false);
            w.EndArg();
        }

        // What a reader needs to read the rest: the clock's rate, the moment it started in UTC, and
        // the settings in force.
        void WriteArmRow(long long qpc, const FILETIME& ft)
        {
            LARGE_INTEGER freq; QueryPerformanceFrequency(&freq);
            SYSTEMTIME st; FileTimeToSystemTime(&ft, &st);
            const unsigned long long ticks = (static_cast<unsigned long long>(ft.dwHighDateTime) << 32) | ft.dwLowDateTime;
            char utc[48];
            _snprintf_s(utc, _TRUNCATE, "%04u-%02u-%02uT%02u:%02u:%02u.%07lluZ", st.wYear, st.wMonth, st.wDay,
                        st.wHour, st.wMinute, st.wSecond, ticks % 10000000ull);
            core::TextBuf buf = {};
            core::WithValueWriter(buf, [&](core::ValueWriter& w)
            {
                w.BeginArgs();
                Named(w, "qpcFrequency", static_cast<long long>(freq.QuadPart));
                Named(w, "utc", std::string(utc));
                Named(w, "pid", static_cast<long long>(GetCurrentProcessId()));
                for (const auto& kv : settings::Take())
                {
                    std::string key = kv.first;
                    for (char& c : key) if (c == ' ') c = '_';
                    Named(w, key.c_str(), kv.second);
                }
                w.EndArgs();
            });
            appevents::WriteRow("XRayXL", "arm", "", buf.Text(), qpc, true);
            buf.Release();
        }

        // The session's totals; a trace without this row stopped with Excel.
        void WriteDisarmRow()
        {
            const vba::Totals vt = g_vbaRan ? vba::ReadTotals() : vba::Totals{};
            const appevents::Totals et = g_eventsRan ? appevents::ReadTotals() : appevents::Totals{};
            core::TextBuf buf = {};
            core::WithValueWriter(buf, [&](core::ValueWriter& w)
            {
                w.BeginArgs();
                Named(w, "rowsDropped", emit::csv::RingDrops());
                Named(w, "pauses", emit::csv::RingPauses());
                Named(w, "valuesRefused", core::ValuesRefused());
                Named(w, "vbaHookFaults", static_cast<long long>(vt.hookFaults));
                Named(w, "vbaStoodDown", vt.tripped);
                Named(w, "stackGrowFailures", static_cast<long long>(vt.stackGrowFailures) + xll::StackGrowFailures());
                Named(w, "xllExitsDropped", xll::ExitsDropped());
                Named(w, "xllHookFaults", xll::RecorderFaults());
                Named(w, "eventsRecorded", et.recorded + et.unknown);
                Named(w, "eventsLost", et.faults);
                w.EndArgs();
            });
            appevents::WriteRow("XRayXL", "disarm", "", buf.Text(), core::QpcNow(), true);
            buf.Release();
        }

        void CloseSession()
        {
            if (!g_sessionOpen) return;
            WriteDisarmRow();
            emit::csv::Close();
            g_sessionOpen = false;
            std::ostringstream d;
            d << "disarmed -- " << emit::csv::RowsWritten() << " row(s) written";
            core::Log::Note(d.str());
        }

        void StartEvents()
        {
            using namespace core::events;
            const Mask chosen = GetSelected();
            std::string why;
            if (!appevents::Start(chosen, why))
            {
                core::Log::Warning("events: not recorded -- " + why);
                return;
            }
            g_eventsRan = true;
            bool known = false;
            const Mask have = appevents::Available(known);
            char list[2048];
            DescribeSelection(chosen, have, list, sizeof(list));
            std::string line = std::string("events: recording ") + list;
            std::string missing;
            for (int i = 0; i < Count(); ++i)
                if (((chosen >> i) & 1) && !((have >> i) & 1))
                    missing += (missing.empty() ? "" : ",") + std::string(At(i).name);
            if (!missing.empty()) line += "; not in this Excel: " + missing;
            if (!known) line += " (this Excel's event list could not be read)";
            core::Log::Note(line);
        }
    }

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

        const long long armQpc = core::QpcNow();
        FILETIME armTime; GetSystemTimePreciseAsFileTime(&armTime);
        g_sessionOpen = emit::csv::Open(core::modes::GetBufferBytes(), core::modes::GetPauseOnFull(),
                                        core::modes::GetFormat(), core::modes::BreaksColumn());
        // The arm row names the events chosen against those this Excel has, so they are read first.
        bool eventsKnown = false;
        appevents::Available(eventsKnown);
        if (!eventsKnown) appevents::Probe(eventsKnown);
        if (g_sessionOpen) WriteArmRow(armQpc, armTime);

        g_vbaRan = g_eventsRan = false;
        core::Log::Note(vba::ArmCounting());
        g_vbaRan = vba::IsVbaArmed();
        core::ResetValuesRefused();
        xll::ArmReport r = xll::Arm();
        if (IsArmed()) StartEvents();
        else           CloseSession();
        // Armed state is mirrored by the ribbon, which may not be what asked.
        core::NotifyStateChanged();
        return r;
    }

    void Disarm(bool shuttingDown)
    {
        // First, so no event row lands after the disarm row. Always, even shutting down: Excel is
        // still alive, and a subscription left behind would call into an unloaded add-in.
        appevents::Stop();

        // The file carries no drop marker, so this line, the disarm row, the Disarm return value
        // and the status summary are the only places loss is reported.
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
        CloseSession();

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
