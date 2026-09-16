#include "xlltrace.h"
#include "xllhook.h"
#include "xllregistry.h"
#include "xllregister_watch.h"
#include "xlltypeplan.h"
#include "emit/csv.h"
#include "core/excel_api.h"
#include "core/log.h"
#include "core/crashlog.h"
#include "core/tracemodes.h"
#include "core/clock.h"
#include "core/text.h"

#include <windows.h>
#include <sstream>
#include <cstdio>

namespace xll
{

    namespace
    {
        // Read by the watch's worker thread as well as the main thread.
        volatile LONG g_isArmed = 0;
        bool Armed() { return InterlockedCompareExchange(&g_isArmed, 0, 0) != 0; }
        void SetArmed(bool on) { InterlockedExchange(&g_isArmed, on ? 1 : 0); }

        void NarrowInto(const std::wstring& w, char* dst, int cap)
        {
            core::NarrowUtf8(w.c_str(), static_cast<int>(w.size()), dst, cap);
        }

        const wchar_t* Leaf(const std::wstring& path)
        {
            const size_t i = path.find_last_of(L'\\');
            return (i == std::wstring::npos) ? path.c_str() : path.c_str() + i + 1;
        }

        HMODULE OwnModule()
        {
            HMODULE h = nullptr;
            GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS |
                               GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                               reinterpret_cast<LPCWSTR>(&OwnModule), &h);
            return h;
        }

        using core::QpcMicros;

        // THE ADDRESS EXCEL CALLS, from a registration's module and export
        // name -- the one step Arm, the late arm and the coverage check share.
        // The module is tried as Excel reported it and then by its leaf, since
        // the loaded name is not always the registered path. Our own module is
        // never a target: its functions are the one add-in whose internals we
        // control, exactly the population that hides the failures worth
        // finding. Declines are counted into `count` when it is given; the
        // coverage check passes nullptr, since it is asking, not arming.
        void* ResolveExport(const std::wstring& module, const std::wstring& procedure,
                            Declines* count)
        {
            HMODULE h = GetModuleHandleW(module.c_str());
            if (h == nullptr) h = GetModuleHandleW(Leaf(module));
            if (h == nullptr)     { if (count) count->moduleNotLoaded++; return nullptr; }
            if (h == OwnModule()) { if (count) count->ownModule++;       return nullptr; }
            char proc[256];      // Excel's 255-character limit; a longer name would narrow to empty
            NarrowInto(procedure, proc, sizeof(proc));
            void* addr = reinterpret_cast<void*>(GetProcAddress(h, proc));
            if (addr == nullptr && count) count->procNotFound++;
            return addr;
        }

        // ---- ONE EXPORT, HOOKED ------------------------------------------
        //
        // The steps Arm() and ArmLate() share once ResolveExport has given
        // them an address. What differs stays with the caller: Arm times the
        // work, ArmLate skips what is already hooked, and the display name is
        // the caller's too -- Arm asks Excel only once a function has been
        // hooked, ArmLate was handed it by xlfRegister.
        //
        // Returns the Target, or nullptr with `why` set and the decline counted
        // -- typetext and stub space here, the detour inside Install().
        struct HookTiming { long long parseUs = 0, installUs = 0; };

        Target* HookExport(void* addr, const std::wstring& procedure, const std::wstring& typeText,
                           const wchar_t* moduleLeaf, std::string& why, HookTiming* tm)
        {
            // Built ONCE; the hot path only reads it. An unparsable type string
            // means declining the function, because tracing it would mean
            // reading arguments from guessed registers.
            const long long tParse = tm ? QpcMicros() : 0;
            Plan plan = Parse(typeText.c_str());
            if (tm) tm->parseUs += QpcMicros() - tParse;
            if (!plan.ok)
            {
                DeclineCounts().typeTextUnparsed++;
                char narrowType[128];
                NarrowInto(typeText, narrowType, sizeof(narrowType));
                char b[192];
                _snprintf_s(b, _TRUNCATE, "typeText '%s' bad code '%c'",
                            narrowType, plan.firstBadCode ? plan.firstBadCode : '?');
                why = b;
                return nullptr;
            }

            // An export hooked in an earlier session gets its own slot back.
            Target* t = FindRetired(addr);
            const bool reused = (t != nullptr);
            if (!reused) t = Allocate();
            if (t == nullptr) { DeclineCounts().stubSpaceExhausted++; why = "no stub space"; return nullptr; }
            if (reused) t->calls = 0;

            t->plan = plan;
            t->argCount = plan.slotCount;   // the thunk reads this, not the plan
            t->exportAddr = addr;
            NarrowInto(procedure, t->procName, sizeof(t->procName));
            NarrowInto(moduleLeaf, t->module, sizeof(t->module));

            const long long tIns = tm ? QpcMicros() : 0;
            const bool installed = Install(t, why);
            if (tm) tm->installUs += QpcMicros() - tIns;
            // Free a fresh slot, or FindTarget would report this export as covered.
            if (!installed) { if (!reused) Release(t); return nullptr; }
            return t;
        }
    }

    // Hooks what registered after arming, on the watch's worker thread.
    // Already-hooked exports are skipped; MinHook would refuse a second detour.
    int ArmLate(const regwatch::Captured* caps, int count)
    {
        if (!Armed() || caps == nullptr) return 0;

        // What this batch hooked, so a failed apply can withdraw just that.
        std::vector<Target*> batch;

        int hooked = 0, skipped = 0, declinedByExcel = 0;
        for (int i = 0; i < count; i++)
        {
            const regwatch::Captured& c = caps[i];
            if (c.module[0] == 0 || c.procedure[0] == 0) continue;

            // A REGISTRATION EXCEL REFUSED IS NOT A FUNCTION. Excel declines
            // any registration attempted DURING a calculation, and the watch
            // still sees the attempt -- hooking it would patch a function no
            // formula can name, then report calls under a name Excel never
            // accepted. The id alone cannot say so: an add-in may ask for no result.
            if (c.refused) { declinedByExcel++; continue; }

            void* addr = ResolveExport(c.module, c.procedure, &DeclineCounts());
            if (addr == nullptr) continue;
            if (Target* live = FindTarget(addr))
            {
                // Hooked already, unless the add-in was unloaded and loaded again at this address.
                std::string why;
                const Reloaded r = RepatchIfReloaded(live, why);
                if (r == Reloaded::Repatched) { batch.push_back(live); hooked++; }
                else if (r == Reloaded::Refused) core::Log::Note("late arm: an export was not hooked again -- " + why);
                else skipped++;
                continue;
            }

            std::string why;
            Target* t = HookExport(addr, c.procedure, c.typeText, Leaf(c.module), why, nullptr);
            if (t == nullptr) continue;
            NarrowInto((c.functionText[0] != 0) ? c.functionText : c.procedure,
                       t->name, sizeof(t->name));
            Publish(t);
            batch.push_back(t);
            hooked++;
        }

        if (hooked > 0)
        {
            std::string why;
            if (!ApplyQueued(why))
            {
                // Keep earlier hooks, but withdraw this batch: its queued enables would
                // otherwise be switched on by the next apply.
                const int abandoned = Withdraw(batch.data(), static_cast<int>(batch.size()));
                core::Log::Note("late arm: could not apply -- " + why +
                                "; the functions registered after arming stay untraced ("
                                + std::to_string(abandoned) + " withdrawn)");
                return 0;
            }
            std::ostringstream m;
            m << "late arm: [tid " << GetCurrentThreadId() << "] hooked "
              << hooked << " function(s) registered after arming"
              << (skipped > 0 ? (" (" + std::to_string(skipped) + " already hooked)") : "")
              << (declinedByExcel > 0
                      ? (" (" + std::to_string(declinedByExcel) + " Excel refused)") : "");
            core::Log::Note(m.str());
        }
        else if (declinedByExcel > 0)
        {
            std::ostringstream m;
            m << "late arm: " << declinedByExcel
              << " registration(s) were refused by Excel and correctly not hooked";
            core::Log::Note(m.str());
        }
        return hooked;
    }

    // WHERE THE TIME WENT, every arm: a total cannot be acted on, a split can.
    // One freeze covers the whole batch, so everything except the apply scales
    // with the function count. Consumes the cost accumulators.
    void ReportArmCost(long long armT0, long long enumUs, long long procUs,
                       long long applyUs, const HookTiming& tm, std::ostringstream& log)
    {
        const ResolveCost rc = TakeResolveCost();
        const long long totalUs = QpcMicros() - armT0;
        char b[320];
        _snprintf_s(b, _TRUNCATE,
                    "arm cost: [tid %lu] total %lldms = enumerate %lldms + name resolution %lldms "
                    "(xlfRegisterId %lldms + xlfGetDef %lldms over %d calls, %d resolved) "
                    "+ hook install %lldms",
                    GetCurrentThreadId(), totalUs / 1000, enumUs / 1000,
                    (rc.regIdUs + rc.getDefUs) / 1000,
                    rc.regIdUs / 1000, rc.getDefUs / 1000, rc.calls, rc.resolved,
                    tm.installUs / 1000);
        core::Log::Note(b);
        log << "  " << b << "\n";

        const InstallCost ic = TakeInstallCost();
        char b2[320];
        _snprintf_s(b2, _TRUNCATE,
                    "arm detail: %d function(s) -- resolve exports %lldus, "
                    "parse %lldus, our stub %lldus, MH_CreateHook %lldus, "
                    "MH_Queue %lldus, ONE MH_ApplyQueued %lldus",
                    ic.calls, procUs, tm.parseUs, ic.stubUs, ic.createUs,
                    ic.queueUs, applyUs);
        core::Log::Note(b2);
        log << "  " << b2 << "\n";
    }

    // WATCH FOR LATE REGISTRATIONS -- ArmLate, off a worker thread. The
    // cost of sitting on every add-in's C API call was measured before anything
    // was built on it: undetectable across 96,002 intercepted calls. A failure
    // to install is NOT a failure to arm; it costs late registrations, which
    // UnhookedRegistrations still counts at disarm.
    void InstallRegisterWatch()
    {
        // An A/B toggle, so the watch's cost can be measured rather than
        // argued about. The variable exists to take the watch AWAY for a
        // control run; unset means on.
        wchar_t off[8]{};
        const bool disabled =
            (GetEnvironmentVariableW(L"XRAYXL_NOREGWATCH", off, 8) > 0 && off[0] == L'1');

        std::string watchWhy;
        if (disabled)
            core::Log::Note("register watch: DISABLED by XRAYXL_NOREGWATCH "
                            "(control run; late registrations will not be hooked)");
        else if (regwatch::Install(&ArmLate, watchWhy))
            core::Log::Note("register watch: installed on MdCallBack12");
        else
            core::Log::Note("register watch: NOT installed -- " + watchWhy +
                            " (late registrations stay invisible; they are still"
                            " counted at disarm)");
    }

    ArmReport Arm()
    {
        ArmReport rep;
        std::ostringstream log;

        if (Armed()) { rep.detail = "already armed"; return rep; }

        ResetDeclines();

        // DEPTH=OFF means NOT HOOKED -- no enumeration, no detours, nothing
        // patched -- never hooked-but-silent.
        SetCapture(core::modes::GetArgs(core::modes::Source::Xll),
                   core::modes::GetRetVal(core::modes::Source::Xll));
        SetTopOnly(core::modes::GetDepth(core::modes::Source::Xll) == core::modes::Depth::Top);

        if (!core::modes::XllEnabled())
        {
            rep.detail = "XLL tracing: OFF -- nothing hooked "
                         "(XRayXL_SetTraceParam(\"XLL\",\"DEPTH\",\"ALL\") to turn it on)";
            core::Log::Note(rep.detail);
            return rep;
        }

        (void)TakeResolveCost();          // discard anything from a previous arm
        (void)TakeInstallCost();
        long long procUs = 0, applyUs = 0;
        HookTiming tm;
        const long long armT0 = QpcMicros();

        std::vector<Registration> regs;
        std::string enumLog;
        if (!Enumerate(regs, enumLog))
        {
            rep.detail = "could not read the registration table: " + enumLog;
            core::Log::Note(rep.detail);
            return rep;
        }
        log << enumLog;
        const long long enumUs = QpcMicros() - armT0;
        rep.registered = static_cast<int>(regs.size());
        int registeredAgain = 0;

        for (auto& r : regs)
        {
            const long long tProc = QpcMicros();
            void* addr = ResolveExport(r.module, r.procedure, &DeclineCounts());
            procUs += QpcMicros() - tProc;
            if (addr == nullptr) { rep.declined++; continue; }
            // The same export registered again under another name is hooked already.
            if (FindTarget(addr) != nullptr) { registeredAgain++; continue; }

            std::string why;
            Target* t = HookExport(addr, r.procedure, r.typeText, Leaf(r.module), why, &tm);
            if (t == nullptr)
            {
                rep.declined++;
                char narrowProc[128];
                NarrowInto(r.procedure, narrowProc, sizeof(narrowProc));
                log << "  declined " << narrowProc << " (" << why << ")\n";
                continue;
            }

            // The name a user would recognise, not the export name -- Excel
            // hands it over directly (ResolveFunctionText), so no cache, no
            // calibration, no heap walk. Falls back to the export name rather
            // than leaving the row anonymous, and the arm line reports how many
            // resolved so an unnamed function stays visible.
            const std::wstring real =
                ResolveFunctionText(r.module, r.procedure, r.typeText);
            NarrowInto(real.empty() ? r.procedure : real, t->name, sizeof(t->name));
            Publish(t);
            rep.armed++;
        }

        // ONE PATCH PASS, fail closed: if it does not take, the detours are in
        // an unknown state, so nothing is armed rather than some of it.
        if (rep.armed > 0)
        {
            std::string applyWhy;
            const long long applyT0 = QpcMicros();
            const bool applied = ApplyQueued(applyWhy);
            applyUs = QpcMicros() - applyT0;
            tm.installUs += applyUs;
            if (!applied)
            {
                DisableAll();
                rep.armed = 0;
                rep.detail = "could not apply the queued detours: " + applyWhy;
                core::Log::Note(rep.detail);
                log << "  " << rep.detail << "\n";
            }
        }

        // Two QPC reads per function against four calls INTO Excel -- cheap
        // enough to leave in.
        ReportArmCost(armT0, enumUs, procUs, applyUs, tm, log);

        InstallRegisterWatch();

        // The VBA side was armed by the session before this ran, so this early
        // return costs it nothing: a workbook with VBA and no registered XLL
        // functions -- the ordinary VBA-only case -- once traced nothing here.
        if (rep.armed == 0)
        {
            DisableAll();
            // Disarm does nothing when nothing is armed, so remove the watch here.
            regwatch::Remove();
            rep.detail = "nothing armed" + std::string(log.str());
            core::Log::Note(rep.detail);
            return rep;
        }

        // Only once something is actually hooked, so an empty file cannot be
        // mistaken for a silent session. Open is idempotent, so a trace the VBA
        // side already opened this arm is kept rather than reopened.
        if (!emit::csv::Open(core::modes::GetBufferBytes(), core::modes::GetPauseOnFull()))
        {
            DisableAll();
            regwatch::Remove();     // as above: nothing is armed, so nothing may stay hooked
            rep.armed = 0;
            rep.detail = "could not open the trace file";
            core::Log::Note(rep.detail);
            return rep;
        }

        SetArmed(true);

        const Declines& d = DeclineCounts();
        std::ostringstream s;
        s << "armed " << rep.armed << " of " << rep.registered
          << " registered; declined " << rep.declined
          << " [module " << d.moduleNotLoaded
          << ", proc " << d.procNotFound
          << ", typetext " << d.typeTextUnparsed
          << ", ours " << d.ownModule
          << ", detour " << d.detourFailed
          << ", space " << d.stubSpaceExhausted << "]";
        if (registeredAgain > 0) s << "; " << registeredAgain << " registered again under another name";
        rep.detail = s.str();
        core::Log::Note(rep.detail);
        core::crashlog::Note(rep.detail.c_str());
        log << "  " << rep.detail << "\n";
        return rep;
    }

    // How many registrations exist NOW that we are not hooking. Coverage is
    // bounded by enumerating once at Arm, which is an acceptable design
    // property; being SILENT about it is not, so it is asked at Disarm and
    // reported.
    int UnhookedRegistrations()
    {
        std::vector<Registration> regs;
        std::string ignored;
        if (!Enumerate(regs, ignored)) return -1;      // could not tell

        int missing = 0;
        for (auto& r : regs)
        {
            void* addr = ResolveExport(r.module, r.procedure, nullptr);   // asking, not arming
            if (addr != nullptr && FindTarget(addr) == nullptr) missing++;
        }
        return missing;
    }

    bool IsArmed() { return Armed(); }

    void Disarm(bool shuttingDown)
    {
        // Idempotent and safe when never armed: it arrives from both
        // Application.Run("XRayXL_Disarm") and xlAutoClose, either first.
        if (!Armed()) return;

        // Stop the watch first, so ArmLate cannot change the target table while it is read.
        regwatch::Remove();

        // BEFORE unhooking, while the target list is still populated. Skipped
        // during shutdown, when the object model is neither safe nor useful.
        const int missed = shuttingDown ? -1 : UnhookedRegistrations();

        // Beside UnhookedRegistrations, which counts the same hole a different way.
        {
            const regwatch::Stats st = regwatch::Snapshot();
            std::ostringstream w;
            w << "register watch: " << st.calls << " C API call(s) seen, "
              << st.registers << " were xlfRegister, " << st.hooked << " hooked, "
              << st.declinedCut << " declined (a field cut short), "
              << st.faults << " fault(s)";
            core::Log::Note(w.str());
        }
        SetArmed(false);          // then, so a re-entrant call is a no-op
        DisableAll();
        // CLOSE IS THE FENCE, so the row count is read AFTER it -- in
        // ring mode the rows are not all written until the drain has run. A
        // quiet trace can be told from a broken one only if the count is
        // printed.
        emit::csv::Close();
        const long long rows = emit::csv::RowsWritten();

        // An entry without its exit reads as a hang. Non-zero only: a clean
        // session says nothing.
        const long long droppedExits = ExitsDropped();
        if (droppedExits > 0)
        {
            std::ostringstream x;
            x << "disarm -- WARNING: " << droppedExits
              << " exit row(s) were dropped because the recorder was re-entered."
                 " Their entry rows have no exit and will read as calls that"
                 " never returned.";
            core::Log::Note(x.str());
        }

        if (const long long f = RecorderFaults())
            core::Log::Warning("disarm -- " + std::to_string(f) + " XLL row(s) lost to faults while decoding");
        if (const long long r = FramesResynced())
            core::Log::Note("disarm -- " + std::to_string(r) + " XLL frame(s) closed after an exception unwound past them");

        if (missed > 0)
        {
            std::ostringstream m;
            m << "disarmed -- WARNING: " << missed
              << " registered function(s) were never hooked. They registered after"
                 " arming, so nothing about them was traced. Arm again to include them.";
            core::Log::Note(m.str());
        }
        else
        {
            std::ostringstream d;
            d << "disarmed -- " << rows << " row(s) written";
            core::Log::Note(d.str());
        }
    }
}
