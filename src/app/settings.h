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
    // (name, value) pairs in a fixed order, e.g. ("XLL DEPTH", "ALL").
    using Snapshot = std::vector<std::pair<std::string, std::string>>;

    Snapshot Take();

    // "XLL DEPTH=ALL, XLL ARGS=TRUE, ..."
    std::string List(const Snapshot& s);

    // "XLL ARGS TRUE -> FALSE; LOGLEVEL INFO -> DEBUG", or empty when nothing differs.
    std::string Changes(const Snapshot& before, const Snapshot& after);
}
}
