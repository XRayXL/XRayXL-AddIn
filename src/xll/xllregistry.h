#pragma once
#include <string>
#include <vector>

// Enumerating what Excel has registered, by asking the documented API. No stack
// walk, no .pdata, no candidate scoring, no calibration -- so nothing here
// depends on a derived address, and the hazards of stack walking do not arise.

namespace xll
{
    struct Registration
    {
        std::wstring module;      // full path as Excel reports it
        std::wstring procedure;   // pxProcedure -- the EXPORT name
        std::wstring typeText;
    };

    // Application.RegisteredFunctions, including for add-ins whose exports are
    // a jump table. NOT GET.WORKSPACE(44), which returns the paths of loaded
    // add-ins rather than the registration table.
    bool Enumerate(std::vector<Registration>& out, std::string& log);

    // WHERE ARMING GOES. A total cannot say which round-trip is expensive, so
    // the two calls are timed separately and reported at every arm.
    struct ResolveCost
    {
        long long regIdUs  = 0;   // xlfRegisterId, plus its xlFree
        long long getDefUs = 0;   // xlfGetDef, plus its xlFree
        int       calls    = 0;
        int       resolved = 0;   // names Excel gave us; the rest fall back
        // The first refusal, kept: without it a name that fell back says only that it did,
        // and xlretInvXlfn -- the C API refusing outside a macro context -- reads as silence.
        int       firstFailFn = 0;   // xlfRegisterId or xlfGetDef
        int       firstFailRc = 0;   // its xlret*, 0 if none failed
    };
    const char* XlRetName(int rc);
    ResolveCost TakeResolveCost();   // reads and resets

    // pxProcedure is the export and pxFunctionText is what goes in a cell; an add-in that
    // generates its exports makes them differ. Resolved inside the C API: xlfRegisterId, then
    // xlfGetDef. Empty if the chain does not resolve, and the caller falls back to the export
    // name.
    std::wstring ResolveFunctionText(const std::wstring& module,
                                     const std::wstring& procedure,
                                     const std::wstring& typeText);
}
