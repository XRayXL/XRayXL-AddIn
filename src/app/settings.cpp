#include "settings.h"
#include "paramparse.h"
#include "core/log.h"
#include "appevents.h"
#include "core/eventlist.h"
#include "core/tracemodes.h"

namespace app
{
namespace settings
{
    namespace
    {
        const char* TrueFalse(bool v) { return v ? "TRUE" : "FALSE"; }

        std::string BufferText(std::size_t bytes)
        {
            wchar_t w[32];
            params::FormatBufferW(bytes, w, 32);
            std::string out;
            for (const wchar_t* p = w; *p; ++p) out += static_cast<char>(*p);   // ASCII only
            return out;
        }
    }

    Snapshot Take()
    {
        using namespace core::modes;
        Snapshot s;
        for (Source src : { Source::Xll, Source::Vba })
        {
            const std::string pre = std::string(SourceName(src)) + " ";
            s.push_back({ pre + "DEPTH",  DepthName(GetDepth(src)) });
            s.push_back({ pre + "ARGS",   TrueFalse(GetArgs(src)),   false });
            s.push_back({ pre + "RETVAL", TrueFalse(GetRetVal(src)), false });
        }
        s.push_back({ "VBA OBJECTS",     TrueFalse(GetObjects(Source::Vba)), false });   // VBA only
        s.push_back({ "VBA BREAKPOINTS", TrueFalse(GetBreakpoints(Source::Vba)), false });
        s.push_back({ "BUFFERSIZE",      BufferText(GetBufferBytes()) });
        s.push_back({ "BUFFERWHENFULL",  GetPauseOnFull() ? "PAUSE" : "DROP" });
        s.push_back({ "FORMAT",          FormatName(GetFormat()) });
        {
            // Against the events this Excel has, as the dialog and the arm log read it.
            bool known = false;
            char ev[2048];
            core::events::DescribeSelection(core::events::GetSelected(), appevents::Available(known), ev, sizeof(ev));
            s.push_back({ "EVENTS", ev });
        }
        s.push_back({ "LOGLEVEL", core::Log::LevelName(core::Log::GetLevel()) });
        return s;
    }

    std::string List(const Snapshot& s)
    {
        std::string out;
        for (const auto& kv : s)
        {
            if (!out.empty()) out += ", ";
            out += kv.name + "=" + kv.value;
        }
        return out;
    }

    std::string Changes(const Snapshot& before, const Snapshot& after)
    {
        std::string out;
        for (const auto& now : after)
            for (const auto& was : before)
                if (was.name == now.name && was.value != now.value)
                {
                    if (!out.empty()) out += "; ";
                    out += now.name + " " + was.value + " -> " + now.value;
                }
        return out;
    }
}
}
