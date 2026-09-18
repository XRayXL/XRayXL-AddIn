// Arm and disarm as XLL commands, reachable from Application.Run without a window, focus or a COM
// object of ours. Commands (macro type 2), not functions: a worksheet function runs where most of
// the Excel API is refused, and arming reads the object model.
//
// Every export below is declared in XRayXL.def, and only there.
#include "core/tracemodes.h"
#include "session.h"
#include "exports.h"
#include "paramparse.h"
#include "core/caller.h"
#include "core/crashlog.h"
#include "core/excel_api.h"
#include "core/log.h"
#include "ui/optionsdlg.h"
#include "emit/csv.h"
#include "xlcall.h"

#include <windows.h>
#include <cstdio>
#include <cwchar>
#include <string>

using namespace app;
using namespace app::params;

namespace
{
    // SEH cannot share a frame with anything that needs unwinding (C2712) and
    // ArmReport holds a std::string, so the __try wrapper below stays a leaf.
    int ArmNoUnwind()
    {
        xll::ArmReport r = app::Arm();
        core::crashlog::Note(r.detail.c_str());
        return 1;
    }

    // Disarm, then report how many rows the run DROPPED: the file carries
    // no `gap` marker, so the count comes back here, and on the status summary
    // and the disarm log line. It survives Close, so it is read after disarming.
    int DisarmReturningDrops()
    {
        app::Disarm();
        return static_cast<int>(emit::csv::RingDrops());
    }

    // A fault in here is a LOGGED bug and Excel survives it. A function
    // pointer needs no unwinding; `onFault` is what the caller gets back for a
    // contained fault.
    int RunGuarded(int (*body)(), int onFault)
    {
        __try
        {
            return body();
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            // No destructor ran in this unwind, so a lock the body held is still held.
            core::Log::ReleaseHeldByThisThread();
            emit::csv::ReleaseHeldByThisThread();
            core::crashlog::Note("XRayXL command FAULTED -- contained; Excel survives");
            return onFault;
        }
    }
}

extern "C" int __stdcall XRayXL_Arm(void)
{
    core::Log::Note("command: XRayXL_Arm");
    return RunGuarded(ArmNoUnwind, 0);
}

// Returns rows DROPPED (0 = none, -1 = a contained fault), so
// `dropped = Application.Run("XRayXL_Disarm")` tells the caller whether the
// trace is complete.
extern "C" int __stdcall XRayXL_Disarm(void)
{
    core::Log::Note("command: XRayXL_Disarm");
    return RunGuarded(DisarmReturningDrops, -1);
}

// ---- XRayXL_Options: the Options dialog as a command, so a macro or a test can open it ----
extern "C" int __stdcall XRayXL_Options(void)
{
    core::Log::Note("command: XRayXL_Options");
    return RunGuarded([] { ui::options::Show(nullptr); return 1; }, 0);
}

// ---- what the worksheet-callable exports share (exports.h) ----------------
namespace app
{
    namespace
    {
        XLOPER12 g_setterFault{};   // handed back if a body faulted
    }

    LPXLOPER12 FaultOper()
    {
        g_setterFault.xltype = xltypeErr;
        g_setterFault.val.err = xlerrValue;
        return &g_setterFault;
    }

    LPXLOPER12 EchoStr(const wchar_t* text)
    {
        static core::PascalStr s_str;
        s_str = core::MakeStr(text);
        return &s_str.oper;
    }

    bool CalledFromCell()
    {
        core::Caller who;      // NOT {} -- ReadCaller fills it on every path
        core::ReadCaller(who);
        return who.isCell;
    }

    bool AnythingArmed() { return IsArmed(); }
}

// AN INSTRUMENT: faults on demand at an address in no module, so the vectored
// crash capture can be PROVEN to work rather than assumed to. EXPECTED TO
// CRASH EXCEL -- that is the measurement. Call it with "EXEC".
namespace
{
    // Split from its wrapper, like every other entry point here: __try may not
    // share a frame with anything that needs object unwinding.
    LPXLOPER12 FaultProbeBody(LPXLOPER12 arg)
    {
        // 'EXEC' jumps to an address in no module -- the exact signature the
        // vectored handler exists for. An exception merely RAISED inside
        // this add-in has its IP in a module, which the handler correctly
        // declines, so it cannot validate anything.
        if (arg && ArgType(arg) == xltypeStr)
        {
            wchar_t b[16]; ReadUpper(arg, b, 16);
            // "DUMPEXEC" turns the opt-in dump on first, so the test can
            // prove BOTH paths: off by default, and written when asked.
            // Fault while holding a lock, so containment can be shown to release it.
            if (!wcscmp(b, L"LOGLOCK")) core::Log::FaultWhileLockedForProbe();
            if (!wcscmp(b, L"CSVLOCK")) emit::csv::FaultWhileLockedForProbe();
            if (!wcscmp(b, L"DUMPEXEC")) core::crashlog::SetDumpEnabled(true);
            if (!wcscmp(b, L"EXEC") || !wcscmp(b, L"DUMPEXEC"))
            {
                core::Log::Note("INJECT: calling through a non-module address");
                typedef void (*Nowhere)();
                // Committed but not executable, so the fault is an EXECUTE
                // violation at an address owning no module rather than a wild
                // pointer that might land anywhere.
                // One page for the process: the call never returns to free it.
                static void* page = VirtualAlloc(nullptr, 4096, MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE);
                Nowhere fn = reinterpret_cast<Nowhere>(page);
                fn();
                return EchoStr(L"unreachable");
            }
        }
        return EchoStr(L"XRayXL_FaultProbe: pass \"EXEC\" to fault deliberately");
    }
}

extern "C" LPXLOPER12 __stdcall XRayXL_FaultProbe(LPXLOPER12 arg)
{
    return GuardedOper(nullptr, FaultProbeBody, arg);
}

// XRayXL_IsArmed: TRUE if either source is armed, by the same test the setters use, so a macro can
// ask before it sets instead of parsing a refusal.
namespace
{
    __declspec(thread) XLOPER12 t_isArmed;

    LPXLOPER12 IsArmedBody()
    {
        t_isArmed.xltype = xltypeBool;
        t_isArmed.val.xbool = AnythingArmed() ? 1 : 0;
        return &t_isArmed;
    }
}

extern "C" LPXLOPER12 __stdcall XRayXL_IsArmed(void)
{
    return GuardedOper("XRayXL_IsArmed FAULTED -- contained", IsArmedBody);
}

namespace app
{
    // False if Excel refused the registration; an add-in whose commands did not register otherwise
    // looks like one that loaded fine. A command (type 2) takes nothing and returns an int; a
    // function (type 1) takes arguments and returns an echo.
    enum class Reg { Command, Function };
    static bool Register(const std::wstring& dll, const wchar_t* proc, const wchar_t* name,
                         Reg kind, const wchar_t* type = L"QQ", const wchar_t* argNames = L"")
    {
        const bool cmd = (kind == Reg::Command);
        core::PascalStr module    = core::MakeStr(dll.c_str());
        core::PascalStr procedure = core::MakeStr(proc);
        core::PascalStr typeText  = core::MakeStr(cmd ? L"J" : type);   // Q per XLOPER12, first is the return
        core::PascalStr funcText  = core::MakeStr(name);
        core::PascalStr argText   = core::MakeStr(cmd ? L"" : argNames);
        core::PascalStr category  = core::MakeStr(L"XRayXL");
        XLOPER12 macroType{}; macroType.xltype = xltypeNum; macroType.val.num = cmd ? 2 : 1;
        XLOPER12 res{};
        int rc = Excel12(xlfRegister, &res, 7,
                         &module.oper, &procedure.oper, &typeText.oper,
                         &funcText.oper, &argText.oper, &macroType, &category.oper);
        bool ok = (rc == xlretSuccess && res.xltype != xltypeErr);
        Excel12(xlFree, nullptr, 1, &res);
        return ok;
    }

    void RegisterCommands()
    {
        const std::wstring dll = core::GetOwnModulePath();
        if (dll.empty()) { core::crashlog::Note("commands: no module path; NOTHING REGISTERED"); return; }
        const bool a = Register(dll, L"XRayXL_Arm",    L"XRayXL_Arm",    Reg::Command);
        const bool d = Register(dll, L"XRayXL_Disarm", L"XRayXL_Disarm", Reg::Command);
        const bool op = Register(dll, L"XRayXL_Options", L"XRayXL_Options", Reg::Command);
        (void)op;
        // Source, Name, Value -- three XLOPER12 arguments and an echo.
        const bool sp = Register(dll, L"XRayXL_SetTraceParam", L"XRayXL_SetTraceParam", Reg::Function, L"QQQQ", L"Source,Name,Value");
        // VOLATILE (the trailing !) throughout below: each of these changes
        // without any argument changing, so a non-volatile cell would sit
        // showing a value that is no longer true.
        const bool gp = Register(dll, L"XRayXL_GetTraceParam", L"XRayXL_GetTraceParam", Reg::Function, L"QQQ!", L"Source,Name");
        // One optional wildcard filter.
        const bool gs = Register(dll, L"XRayXL_GetTraceSummary", L"XRayXL_GetTraceSummary", Reg::Function, L"QQ!", L"Filter");
        // Diagnostic only: it faults on purpose, so it is registered only with XRAYXL_DIAG=1.
        const bool diag = core::modes::DiagEnabled();
        const bool fp = diag && Register(dll, L"XRayXL_FaultProbe", L"XRayXL_FaultProbe", Reg::Function, L"QQ", L"Mode");
        // No arguments.
        const bool ia = Register(dll, L"XRayXL_IsArmed", L"XRayXL_IsArmed", Reg::Function, L"Q!");
        char line[288];
        _snprintf_s(line, _TRUNCATE,
                    "commands registered: XRayXL_Arm=%s XRayXL_Disarm=%s"
                    " XRayXL_SetTraceParam=%s XRayXL_GetTraceParam=%s XRayXL_GetTraceSummary=%s"
                    " XRayXL_FaultProbe=%s XRayXL_IsArmed=%s"
                    "  (call with Application.Run)",
                    a ? "ok" : "FAILED", d ? "ok" : "FAILED",
                    sp ? "ok" : "FAILED", gp ? "ok" : "FAILED", gs ? "ok" : "FAILED",
                    fp ? "ok" : (diag ? "FAILED" : "off"), ia ? "ok" : "FAILED");
        core::crashlog::Note(line);
    }
}
