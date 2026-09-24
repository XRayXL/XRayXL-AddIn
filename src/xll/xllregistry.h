#pragma once
#include <string>
#include <vector>

// What Excel has registered, from the documented API, so nothing depends on a derived address.

namespace xll
{
    struct Registration
    {
        std::wstring module;      // full path as Excel reports it
        std::wstring procedure;   // pxProcedure -- the export name
        std::wstring typeText;
    };

    // Application.RegisteredFunctions, which covers exports that are a jump table. Not
    // GET.WORKSPACE(44): that lists loaded add-in paths, not registrations.
    bool Enumerate(std::vector<Registration>& out, std::string& log);

    // The two round-trips timed separately, because a total cannot say which is expensive.
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

    // pxFunctionText, the name used in a cell, which differs from the export when an add-in
    // generates its exports. Empty if xlfRegisterId then xlfGetDef does not resolve it.
    std::wstring ResolveFunctionText(const std::wstring& module,
                                     const std::wstring& procedure,
                                     const std::wstring& typeText);
}
