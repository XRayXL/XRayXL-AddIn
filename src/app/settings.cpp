#include "settings.h"
#include "paramparse.h"
#include "core/log.h"
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
            s.emplace_back(pre + "DEPTH",  DepthName(GetDepth(src)));
            s.emplace_back(pre + "ARGS",   TrueFalse(GetArgs(src)));
            s.emplace_back(pre + "RETVAL", TrueFalse(GetRetVal(src)));
        }
        s.emplace_back("VBA OBJECTS",    TrueFalse(GetObjects(Source::Vba)));   // VBA only
        s.emplace_back("VBA BREAKPOINTS", TrueFalse(GetBreakpoints(Source::Vba)));
        s.emplace_back("BUFFERSIZE",     BufferText(GetBufferBytes()));
        s.emplace_back("BUFFERWHENFULL", GetPauseOnFull() ? "PAUSE" : "DROP");
        s.emplace_back("FORMAT",         FormatName(GetFormat()));
        s.emplace_back("LOGLEVEL",       core::Log::LevelName(core::Log::GetLevel()));
        return s;
    }

    std::string List(const Snapshot& s)
    {
        std::string out;
        for (const auto& kv : s)
        {
            if (!out.empty()) out += ", ";
            out += kv.first + "=" + kv.second;
        }
        return out;
    }

    std::string Changes(const Snapshot& before, const Snapshot& after)
    {
        std::string out;
        for (const auto& now : after)
            for (const auto& was : before)
                if (was.first == now.first && was.second != now.second)
                {
                    if (!out.empty()) out += "; ";
                    out += now.first + " " + was.second + " -> " + now.second;
                }
        return out;
    }
}
}
