#pragma once
#include <string>

// Hooks functions that register after arming, which the one enumeration at arm would miss.
//
// Every C API call from every XLL goes through one export of EXCEL.EXE, MdCallBack12, so one
// hook there sees every registration.
//
// Captures are batched and a worker patches them a beat later, because MinHook suspends every
// thread for each patch. The worker cannot call Excel, so the capture copies everything the
// patcher needs.

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
