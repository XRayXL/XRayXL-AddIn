// XRayXL -- the XLL vehicle. Everything is reachable without COM; it is also an in-proc COM server
// for the ribbon buttons alone (src/ui/ribbon.cpp), the one part allowed to fail.

#include "core/excel_api.h"
#include "xlcall.h"
#include "core/crashlog.h"
#include "core/log.h"
#include "app/session.h"
#include "app/settings.h"
#include "core/text.h"
#include "emit/csv.h"
#include "vba/vbatrace.h"
#include "xll/xlltrace.h"
#include "ui/ribbon.h"

#include <windows.h>
#include <sstream>
#include <string>

namespace
{
    // The product version, from version.props as three numeric defines, because a quoted-string
    // define does not survive the resource compiler's command line. The 0.0.0 fallback is meant to
    // look broken.
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

    // Set by xlAutoRemove: the add-in dialog is removing us, and no COM callback will follow.
    volatile LONG g_removing = 0;

    // The process id, because that is what a person can match against Task Manager.
    std::wstring SessionSuffix()
    {
        std::wostringstream o;
        o << L"_" << GetCurrentProcessId();
        return o.str();
    }

    // With %TEMP%\XRayXL\inert.on present, xlAutoOpen returns before doing anything at all. If a
    // crash survives that, only the add-in's presence is the cause.
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

// The export list lives in XRayXL.def, and only there: Excel finds these by their plain names,
// which a .def keeps undecorated on x86 too.

// ---- the two COM exports, for the ribbon and nothing else ----------------

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
{
    return ui::ribbon::GetClassObject(rclsid, riid, ppv);
}

// Always S_FALSE: while loaded this module may have detours and dispatch slots in other modules,
// so there is never a safe moment for COM to unload it.
extern "C" HRESULT __stdcall DllCanUnloadNow()
{
    return S_FALSE;
}

namespace app
{
    const char* VersionText() { return kVersionText; }
}

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
        // the version, so copies of different builds can be told apart in the Add-ins dialog
        wchar_t text[48];
        _snwprintf_s(text, _TRUNCATE, L"XRayXL XLL v%hs", kVersionText);
        name = core::MakeStr(text);
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
    // Where this XLL was loaded from, so a support log names the exact binary; at Info, so it
    // shows in a shipped log.
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
    // XLL commands, so Application.Run reaches them with no window and no COM object of ours.
    app::RegisterCommands();
    core::Log::Note("loaded: XLL commands registered.");
    core::Log::Note("settings at start -- " + app::settings::List(app::settings::Take()));

    // last, so a ribbon that cannot load costs the session its buttons and nothing else
    ui::ribbon::Start();

  } catch (...) { core::crashlog::Note("xlAutoOpen: exception contained"); }
    return 1;
}

extern "C" int __stdcall xlAutoRemove()
{
    InterlockedExchange(&g_removing, 1);
  try {
    core::Log::Note("xlAutoRemove: the add-in is being removed; xlAutoClose will tear down");
  } catch (...) { core::crashlog::Note("xlAutoRemove: exception contained"); }
    return 1;
}

extern "C" int __stdcall xlAutoClose()
{
  try {
    core::crashlog::Note("xlAutoClose: entered");

    // Excel calls this before its Save prompt, so the user may still Cancel. When the
    // ribbon is connected, its OnDisconnection tears down once the exit is certain.
    const bool removing = InterlockedExchange(&g_removing, 0) != 0;
    if (!removing && ui::ribbon::Connected())
    {
        core::Log::Note("xlAutoClose: Excel may still cancel; teardown waits for OnDisconnection");
        return 1;
    }
    core::Log::Note(removing ? "xlAutoClose: add-in removed -- disarming and stopping the ribbon"
                             : "xlAutoClose: no ribbon to report the exit -- disarming now");

    // Unhook first: a detour left behind after the module goes points at
    // unmapped code the next time Excel calls that add-in function, which is a
    // crash with our name nowhere near it.
    app::Disarm(true);

    // after Disarm, so nothing is patched into anyone else's module if the disconnect were to fault
    ui::ribbon::Stop();

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
    if (reason == DLL_THREAD_DETACH)
    {
        emit::csv::ReleaseThreadScratch();
        vba::ReleaseThreadState();
        xll::ReleaseThreadState();
    }
    if (reason == DLL_PROCESS_DETACH)
    {
        // Non-null lpReserved means the process is exiting: the other threads were terminated
        // wherever they were, possibly holding the heap lock, so do nothing.
        if (lpReserved != nullptr) return TRUE;

        // A genuine FreeLibrary -- unreachable with the module pinned, and
        // nothing to undo if it happens.
    }
    return TRUE;
}
