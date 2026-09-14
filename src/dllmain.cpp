// XRayXL -- the XLL vehicle.
//
// A PURE XLL: Excel loads this module, calls it and unloads it -- one owner,
// one lifetime. No COM server, no ribbon and no dialog; arming is an XLL
// command, Application.Run("XRayXL_Arm").
//
// The export surface is the xlAuto* entry points plus the commands and
// functions in src/app/commands.cpp.

#include "core/excel_api.h"
#include "xlcall.h"
#include "core/crashlog.h"
#include "core/log.h"
#include "app/session.h"
#include "core/text.h"
#include "emit/csv.h"


#include <windows.h>
#include <sstream>
#include <string>

namespace
{
    // THE PRODUCT VERSION, from version.props via the project file. It arrives
    // as three NUMERIC defines and is assembled here, because a define whose
    // value is a quoted string does not survive the resource compiler's
    // command line -- and both consumers must read the same defines or the
    // resource and the log can disagree about which build this is. The
    // fallback is deliberately not a plausible version: 0.0.0 in a log means
    // the plumbing broke, and that should look broken.
#if !defined(XRAY_VER_MAJOR) || !defined(XRAY_VER_MINOR) || !defined(XRAY_VER_PATCH)
#define XRAY_VER_MAJOR 0
#define XRAY_VER_MINOR 0
#define XRAY_VER_PATCH 0
#endif
#define XRAY_STR2(x) #x
#define XRAY_STR(x)  XRAY_STR2(x)
    constexpr const char* kVersionText =
        XRAY_STR(XRAY_VER_MAJOR) "." XRAY_STR(XRAY_VER_MINOR) "." XRAY_STR(XRAY_VER_PATCH);

    bool g_modulePinned = false;

    // The process id, because that is what a person can match against Task
    // Manager.
    std::wstring SessionSuffix()
    {
        std::wostringstream o;
        o << L"_" << GetCurrentProcessId();
        return o.str();
    }

    // THE BOTTOM RUNG OF THE BISECTION. With %TEMP%\XRayXL\inert.on present,
    // xlAutoOpen returns before the module is pinned, before a log is opened,
    // before anything at all -- an XLL Excel loads that does nothing whatever.
    // If a crash survives that, nothing this add-in DOES is the cause; only its
    // presence.
    bool MarkerPresent(const wchar_t* name)
    {
        wchar_t tmp[MAX_PATH] = {};
        if (!GetTempPathW(MAX_PATH, tmp)) return false;
        wchar_t f[MAX_PATH * 2];
        _snwprintf_s(f, _TRUNCATE, L"%sXRayXL%c%s", tmp, (wchar_t)0x5C, name);
        return GetFileAttributesW(f) != INVALID_FILE_ATTRIBUTES;
    }
    bool Inert() { return MarkerPresent(L"inert.on"); }
}

// THE EXPORT LIST LIVES IN XRayXL.def, AND ONLY THERE.
//
// Not __declspec(dllexport): Excel finds these by GetProcAddress under their
// PLAIN names, and while an extern "C" __stdcall function is undecorated on
// x64, on x86 the same declaration decorates to _xlAutoOpen@0 -- Excel would
// find nothing, silently, in an add-in that otherwise loads fine. A .def names
// the export undecorated on both, which is why it is kept while the build is
// x64-only (32-bit is a stated stretch goal).
//
// One source, so a miss is LOUD: a function absent from the .def is not
// exported at all, and commands.cpp reports its registration as FAILED.

// NO COM SERVER EXPORTS, and so no second lifetime owner for the module.

extern "C" LPXLOPER12 __stdcall xlAddInManagerInfo12(LPXLOPER12 xAction)
{
    static XLOPER12 result{};
    static core::PascalStr name;

    // Code 1 is the add-in's display name; anything else gets an error, which
    // is what the C API expects.
    double action = 0;
    if (xAction != nullptr && xAction->xltype == xltypeNum) action = xAction->val.num;
    else if (xAction != nullptr && xAction->xltype == xltypeInt) action = xAction->val.w;

    if (static_cast<int>(action) == 1)
    {
        name = core::MakeStr(L"XRayXL");
        result = name.oper;
    }
    else
    {
        result.xltype = xltypeErr;
        result.val.err = xlerrValue;
    }
    return &result;
}

extern "C" int __stdcall xlAutoOpen()
{
  try {
    // The bottom rung: loaded by Excel, and doing nothing whatever.
    if (Inert()) return 1;
    // Excel calls this again when the add-in is added again. The commands are
    // registered again; the process-wide setup below happens once.
    static volatile LONG s_setUp = 0;
    if (InterlockedExchange(&s_setUp, 1) != 0)
    {
        core::Log::Note("xlAutoOpen again: commands registered again; the log and crash handlers were already set up");
        app::RegisterCommands();
        return 1;
    }
    // Pinning removes a dilemma rather than managing it: the module cannot be
    // unloaded, so DLL_PROCESS_DETACH can only mean process exit -- where the
    // right thing to do is nothing at all.
    {
        HMODULE self = nullptr;
        BOOL pinned = GetModuleHandleExW(
            GET_MODULE_HANDLE_EX_FLAG_PIN | GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
            reinterpret_cast<LPCWSTR>(&xlAutoOpen), &self);
        g_modulePinned = (pinned != FALSE);
    }

    core::crashlog::Install(core::EnsureAppSubdir(L"Logs") + L"\\XRayXL_crash" + SessionSuffix() + L".txt");
    core::crashlog::InstallVectored(core::EnsureAppSubdir(L"Logs"));
    core::Log::Open(core::EnsureAppSubdir(L"Logs") + L"\\XRayXL" + SessionSuffix() + L".log");
    if (const char* why = core::OutputDirRefused())
        core::Log::Warning(std::string("XRAYXL_OUTPUT_DIR was not used -- ") + why +
                           "; output is under %TEMP%\\XRayXL instead");
    core::crashlog::Note(g_modulePinned ? "module pinned; it cannot be unloaded"
                                  : "module NOT pinned; unload remains a risk");
    // WHERE this XLL was loaded from -- the project build, or a copy -- so a
    // support log names the exact binary. INFO, so it shows in a shipped log.
    {
        const std::wstring xllPath = core::GetOwnModulePath();
        // UTF-8 can need up to 4 bytes per UTF-16 unit.
        char narrow[MAX_PATH * 4]{};
        if (!xllPath.empty())
            core::NarrowUtf8(xllPath.c_str(), -1, narrow, sizeof(narrow));
        char msg[sizeof(narrow) + 96];
        _snprintf_s(msg, _TRUNCATE, "session start: XRayXL %s, pid %lu, xll=%s",
                    kVersionText, GetCurrentProcessId(),
                    narrow[0] ? narrow : "(unknown)");
        core::Log::Info(msg);
    }
    // XLL commands, so Application.Run reaches them from VBA and from any
    // automation client with no window and no COM object of ours in the process
    // (src/app/commands.cpp).
    app::RegisterCommands();
    core::Log::Note("loaded: XLL commands registered; no COM server, no ribbon, no dialog.");

  } catch (...) { core::crashlog::Note("xlAutoOpen: exception contained"); }
    return 1;
}

extern "C" int __stdcall xlAutoClose()
{
  try {
    core::crashlog::Note("xlAutoClose: entered");

    // Unhook FIRST: a detour left behind after the module goes points at
    // unmapped code the next time Excel calls that add-in function, which is a
    // crash with our name nowhere near it.
    app::Disarm(true);

    // Nothing else to bring down: the commands are registered with Excel, and
    // Excel forgets them when it unloads the XLL.
    core::Log::Note("unloaded");
    core::crashlog::Note("xlAutoClose: done");
  } catch (...) { core::crashlog::Note("xlAutoClose: exception contained"); }
    return 1;
}

extern "C" void __stdcall xlAutoFree12(LPXLOPER12) {}

BOOL APIENTRY DllMain(HMODULE, DWORD reason, LPVOID lpReserved)
{
    // A thread that ends normally gives back its row scratch; never at process exit, below.
    if (reason == DLL_THREAD_DETACH) emit::csv::ReleaseThreadScratch();
    if (reason == DLL_PROCESS_DETACH)
    {
        // lpReserved is non-null when the PROCESS is exiting, null for a real
        // FreeLibrary. At process exit every other thread has already been
        // terminated wherever it happened to be, quite possibly holding the heap
        // lock -- so doing anything is both the worst moment for it and
        // pointless.
        if (lpReserved != nullptr) return TRUE;

        // A genuine FreeLibrary -- unreachable with the module pinned, and
        // nothing to undo if it happens.
    }
    return TRUE;
}
