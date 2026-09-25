#pragma once
#include <string>
#include "core/eventlist.h"

// Excel's Application events, recorded as `event` rows while tracing is armed. Everything here runs
// on Excel's main thread: arming, disarming and the events themselves.
namespace app
{
namespace appevents
{
    // At arm, once the trace is open: reads which events this Excel has from its own type library,
    // and subscribes. False, with `why`, when there is no Application to listen to.
    bool Start(core::events::Mask selected, std::string& why);
    void Stop();

    // The catalogue's events this Excel has, as the last Start or Probe read them; `known` is false
    // when the type library could not be read.
    core::events::Mask Available(bool& known);

    // Reads the type library without subscribing, for the Options dialog.
    core::events::Mask Probe(bool& known);

    // "Sh As Object, Target As Range" for catalogue event `k`, from the last Start or Probe.
    bool ParamsText(int k, wchar_t* out, int cap);

    struct Totals
    {
        long long recorded = 0;     // rows written for events the catalogue knows
        long long unknown  = 0;     // rows for events a newer Excel added
        long long faults   = 0;     // events lost to a fault while describing them
    };
    Totals ReadTotals();

    // An `event` row: span, parent and depth 0, since an event is not a call.
    // `keep`: never dropped, for the session's own arm and disarm rows.
    void WriteRow(const char* source, const char* function, const char* callerref,
                  const char* args, long long qpc, bool keep = false);
}
}
