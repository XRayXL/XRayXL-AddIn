#pragma once
#include <string>

// HOOKING FUNCTIONS THAT REGISTER AFTER WE ARMED -- arming enumerates
// Application.RegisteredFunctions ONCE, so an add-in loaded later would never be
// traced.
//
// WHERE TO WATCH: an XLL does not call xlfRegister directly but goes through
// Excel12, which resolves one exported entry point in EXCEL.EXE --
// GetProcAddress(GetModuleHandle(NULL), "MdCallBack12"), at ordinal 46, the same
// one this add-in's own Excel12 uses. Every C API call from every XLL passes
// through it, so one hook on a documented export sees every registration: no
// derivation, no signature, no scanning.
//
// NOT PATCHED ON THE SPOT, because MinHook's patch suspends every thread in the
// process and nothing in the hook knows which registration ends a burst -- so
// inline patching costs one whole-process freeze PER FUNCTION:
//
//     RegisterXLL of a 47-function XLL, inline    2,829 ms
//     RegisterXLL of a 47-function XLL, deferred     31 ms
//
// Both hook all 47. Captures are therefore collected and a worker applies them
// as one batch a beat later. The worker cannot ask Excel anything -- the object
// model is main-thread only and the C API is not callable from an arbitrary
// thread -- so the capture copies everything the patcher will need, and the
// worker touches only GetModuleHandle, GetProcAddress and MinHook.

namespace xll
{
    namespace regwatch
    {
        struct Captured
        {
            wchar_t module[260];
            wchar_t procedure[128];
            // Wide enough for 255 arguments; a truncated type text parses to the wrong arity.
            wchar_t typeText[576];
            wchar_t functionText[128];
            double  id;          // what xlfRegister returned, when the caller asked for a result
            // Excel said no: a failed call, or an error where the id would be. A caller
            // that asked for no result leaves only the return code to go on.
            bool    refused;
        };

        struct Stats
        {
            long long calls     = 0;  // every C API call seen -- the cost figure
            long long registers = 0;  // how many were xlfRegister
            long long hooked    = 0;  // and how many of those were patched
            long long declinedCut = 0; // a field longer than the capture: correctly not hooked
            long long faults    = 0;  // the guard fired -- must stay zero
        };

        // Runs on the worker thread, so it may allocate and take its time.
        typedef int (*ApplyFn)(const Captured* caps, int count);

        // Called from the arming path, never from a traced thread.
        bool Install(ApplyFn apply, std::string& why);
        void Remove();
        Stats Snapshot();
    }
}
