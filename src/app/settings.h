// Every setting at once, named as XRayXL_SetTraceParam names it, for the log: the whole list at
// start and when Options opens, and what an Apply actually changed.
#pragma once
#include <string>
#include <utility>
#include <vector>

namespace app
{
namespace settings
{
    // In a fixed order, e.g. { "XLL DEPTH", "ALL" }. `text` is false for an on/off setting, whose
    // value is TRUE or FALSE.
    struct Setting { std::string name, value; bool text = true; };
    using Snapshot = std::vector<Setting>;

    Snapshot Take();

    // "XLL DEPTH=ALL, XLL ARGS=TRUE, ..."
    std::string List(const Snapshot& s);

    // "XLL ARGS TRUE -> FALSE; LOGLEVEL INFO -> DEBUG", or empty when nothing differs.
    std::string Changes(const Snapshot& before, const Snapshot& after);
}
}
