// XRayXL_SetTraceParam and XRayXL_GetTraceParam. Functions, not commands, because each needs an
// argument and an echo. Both refuse a cell caller, since a formula would reconfigure the tracer on
// every recalc, and both refuse while armed, since the modes are read once at arm.
#include "session.h"
#include "exports.h"
#include "paramparse.h"
#include "xlgrid.h"
#include "core/log.h"
#include "core/ascii.h"
#include "core/tracemodes.h"
#include "xlcall.h"

#include <windows.h>
#include <cstdio>
#include <cwchar>
#include <string>

using namespace app;
using namespace app::params;

namespace
{
    // One setting per call, addressed by (Source, Name), so a call changes only what it names and a
    // new setting is a new name rather than a wider parser.
    //
    // Omitting Source addresses both. An unsupplied XLL argument arrives as xltypeMissing whatever
    // its position, so a leading argument can be omitted without a placeholder.
    using namespace app::params;

    void ApplyOne(core::modes::Source s, core::modes::Param p, core::modes::Depth d, bool on)
    {
        switch (p)
        {
        case core::modes::Param::Depth:  core::modes::SetDepth (s, d);  break;
        case core::modes::Param::Args:   core::modes::SetArgs  (s, on); break;
        case core::modes::Param::RetVal: core::modes::SetRetVal(s, on); break;
        // Named, not `default:`, so a new param cannot silently set another.
        case core::modes::Param::Objects: core::modes::SetObjects(s, on); break;
        }
    }

    void ValueTextW(core::modes::Source s, core::modes::Param p, wchar_t* out, int cap)
    {
        const wchar_t* t;
        switch (p)
        {
        case core::modes::Param::Depth:  t = core::modes::DepthNameW(core::modes::GetDepth(s));  break;
        case core::modes::Param::Args:   t = core::modes::OnOffW(core::modes::GetArgs(s));       break;
        case core::modes::Param::RetVal: t = core::modes::OnOffW(core::modes::GetRetVal(s));     break;
        // OBJECTS is VBA only; TRUE against XLL would promise an effect it has not.
        default: t = (s == core::modes::Source::Xll) ? L"n/a -- VBA only"
                                                     : core::modes::OnOffW(core::modes::GetObjects(s));
                 break;
        }
        wcsncpy_s(out, cap, t, _TRUNCATE);
    }

    // A SOURCE-LESS setting lands its word in whichever slot the caller used
    // -- SetTraceParam("BUFFERSIZE", n) puts it in the Source slot -- so the
    // value is whatever follows the word. nullptr when this is not that word.
    LPXLOPER12 ValueAfterWord(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg,
                              const wchar_t* word)
    {
        if (IsWord(srcArg,  word)) return nameArg;
        if (IsWord(nameArg, word)) return valArg;
        return nullptr;
    }
    bool NamesWord(LPXLOPER12 srcArg, LPXLOPER12 nameArg, const wchar_t* word)
    {
        return IsWord(srcArg, word) || IsWord(nameArg, word);
    }

    // BUFFERSIZE has no source, so the word is accepted in either slot with the value after it.
    // Returns the echo, or nullptr when this was not a BUFFERSIZE call.
    LPXLOPER12 HandleBufferSizeParam(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg)
    {
        LPXLOPER12 bufVal = ValueAfterWord(srcArg, nameArg, valArg, L"BUFFERSIZE");
        if (!bufVal) return nullptr;

        if (AnythingArmed())
            return EchoStr(L"#Err - cannot change settings while armed; XRayXL_Disarm first, then set, then XRayXL_Arm");
        unsigned long long bytes = 0;
        if (!ParseBufferBytes(bufVal, bytes))
            return EchoStr(L"#Err - BUFFERSIZE is a number with optional K/KB or M/MB "
                           L"(bare = MB), 0..240 MB, 0 = synchronous; nothing changed");
        if (bytes != 0 && bytes < kBufMinRing)
            return EchoStr(L"#Err - BUFFERSIZE ring must be at least 16 KB, "
                           L"or 0 for synchronous; nothing changed");
        core::modes::SetBufferBytes(static_cast<std::size_t>(bytes));
        wchar_t sz[32]; FormatBufferW(bytes, sz, 32);
        char line[128];
        _snprintf_s(line, _TRUNCATE, "trace param set: BUFFERSIZE -> %S", sz);
        core::Log::Note(line);
        static wchar_t echo[160];
        if (bytes == 0)
            _snwprintf_s(echo, _TRUNCATE, L"BUFFERSIZE=0 (synchronous, crash-safe; takes effect at next arm)");
        else
            _snwprintf_s(echo, _TRUNCATE, L"BUFFERSIZE=%s (ring, drained off-thread; takes effect at next arm)", sz);
        return EchoStr(echo);
    }

    // BUFFERWHENFULL is source-less too: DROP or PAUSE, only meaningful when
    // BUFFERSIZE=N. Same landing rule as BUFFERSIZE; same return convention.
    LPXLOPER12 HandleBufferWhenFullParam(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg)
    {
        LPXLOPER12 wfVal = ValueAfterWord(srcArg, nameArg, valArg, L"BUFFERWHENFULL");
        if (!wfVal) return nullptr;

        if (AnythingArmed())
            return EchoStr(L"#Err - cannot change settings while armed; XRayXL_Disarm first, then set, then XRayXL_Arm");
        bool pause = false;
        if (!ParseWhenFull(wfVal, pause))
            return EchoStr(L"#Err - BUFFERWHENFULL must be DROP or PAUSE; nothing changed");
        core::modes::SetPauseOnFull(pause);
        char line[96];
        _snprintf_s(line, _TRUNCATE, "trace param set: BUFFERWHENFULL -> %s", pause ? "PAUSE" : "DROP");
        core::Log::Note(line);
        return EchoStr(pause
            ? L"BUFFERWHENFULL=PAUSE (a full ring makes traced threads wait until it is half empty; never loses; takes effect at next arm)"
            : L"BUFFERWHENFULL=DROP (rows dropped when the ring is full, each leaving a hole in the input column; takes effect at next arm)");
    }

    // FORMAT is source-less too: the file's format, CSV or JSONL. Same landing rule and return
    // convention as BUFFERWHENFULL.
    LPXLOPER12 HandleFormatParam(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg)
    {
        LPXLOPER12 v = ValueAfterWord(srcArg, nameArg, valArg, L"FORMAT");
        if (!v) return nullptr;

        if (AnythingArmed())
            return EchoStr(L"#Err - cannot change settings while armed; XRayXL_Disarm first, then set, then XRayXL_Arm");
        core::modes::Format f = core::modes::Format::Csv;
        if (!ParseFormat(v, f))
            return EchoStr(L"#Err - FORMAT must be CSV or JSONL; nothing changed");
        core::modes::SetFormat(f);
        char line[64];
        _snprintf_s(line, _TRUNCATE, "trace param set: FORMAT -> %s", core::modes::FormatName(f));
        core::Log::Note(line);
        return EchoStr(f == core::modes::Format::Jsonl
            ? L"FORMAT=JSONL (one JSON object a line, every value typed; takes effect at next arm)"
            : L"FORMAT=CSV (one row a line, values as text; takes effect at next arm)");
    }

    // LOGLEVEL is source-less and, unlike the trace params, settable AT ANY
    // TIME: it controls the log, not the trace, so it is not refused while
    // armed. Same return convention.
    LPXLOPER12 HandleLogLevelParam(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg)
    {
        LPXLOPER12 v = ValueAfterWord(srcArg, nameArg, valArg, L"LOGLEVEL");
        if (!v) return nullptr;

        core::Log::Level lvl;
        wchar_t wide[16] = {};
        char    word[16] = {};
        if (ArgType(v) == xltypeStr) ReadUpper(v, wide, 16);
        core::NarrowAscii(wide, word, sizeof(word));
        if (!core::Log::LevelFromText(word, lvl))
            return EchoStr(L"#Err - LOGLEVEL must be DEBUG, INFO, WARNING or ERROR; nothing changed");

        core::Log::SetLevel(lvl);
        core::Log::Info(std::string("log level set to ") + core::Log::LevelName(lvl));
        static wchar_t echo[80];
        _snwprintf_s(echo, _TRUNCATE, L"LOGLEVEL=%S (effective now)", core::Log::LevelName(lvl));
        return EchoStr(echo);
    }

    LPXLOPER12 SetTraceParamBody(LPXLOPER12 srcArg, LPXLOPER12 nameArg, LPXLOPER12 valArg)
    {
        core::Log::Note("function: XRayXL_SetTraceParam");
        if (CalledFromCell())
            return EchoStr(L"#Err - call via Application.Run, not from a cell");

        // LOGLEVEL first, because it is settable while armed; then the two
        // source-less words that do refuse. Each returns its own echo or
        // nullptr, falling through to the ordinary Source/Name/Value path.
        if (LPXLOPER12 r = HandleLogLevelParam(srcArg, nameArg, valArg)) return r;
        if (LPXLOPER12 r = HandleBufferSizeParam(srcArg, nameArg, valArg))   return r;
        if (LPXLOPER12 r = HandleBufferWhenFullParam(srcArg, nameArg, valArg)) return r;
        if (LPXLOPER12 r = HandleFormatParam(srcArg, nameArg, valArg))         return r;

        bool both = false; core::modes::Source s = core::modes::Source::Xll;
        if (!ParseSource(srcArg, both, s))
            return EchoStr(L"#Err - Source must be XLL or VBA, or omitted for both; nothing changed");

        core::modes::Param p = core::modes::Param::Depth;
        // With no Source, the settings of the recording as a whole are candidates too.
        if (!ParseParam(nameArg, p))
            return EchoStr(both
                ? L"#Err - Name must be DEPTH, ARGS, RETVAL, OBJECTS, BUFFERSIZE, BUFFERWHENFULL, FORMAT or LOGLEVEL; nothing changed"
                : L"#Err - Name must be DEPTH, ARGS, RETVAL or OBJECTS; nothing changed");

        // OBJECTS is VBA only, so it is refused for XLL and for an omitted Source: "both" cannot mean one.
        if (p == core::modes::Param::Objects && (both || s == core::modes::Source::Xll))
            return EchoStr(L"#Err - OBJECTS Parameter only available for VBA");

        if (AnythingArmed())
            return EchoStr(L"#Err - cannot change settings while armed; XRayXL_Disarm first, then set, then XRayXL_Arm");

        core::modes::Depth d = core::modes::Depth::Off; bool on = false;
        if (p == core::modes::Param::Depth)
        {
            if (!ParseDepth(valArg, d))
                return EchoStr(L"#Err - DEPTH must be OFF, TOP or ALL; nothing changed");
        }
        else
        {
            if (!ParseOnOff(valArg, on))
                return EchoStr(L"#Err - value must be TRUE or FALSE; nothing changed");
        }

        if (both) { ApplyOne(core::modes::Source::Xll, p, d, on); ApplyOne(core::modes::Source::Vba, p, d, on); }
        else      { ApplyOne(s, p, d, on); }

        wchar_t vx[24], vv[24];
        ValueTextW(core::modes::Source::Xll, p, vx, 24);
        ValueTextW(core::modes::Source::Vba, p, vv, 24);
        char line[192];
        _snprintf_s(line, _TRUNCATE, "trace param set: %s %s -> XLL=%S VBA=%S",
                    both ? "BOTH" : core::modes::SourceName(s), core::modes::ParamName(p), vx, vv);
        core::Log::Note(line);

        static wchar_t echo[192];
        if (both)
            _snwprintf_s(echo, _TRUNCATE, L"%s: XLL=%s VBA=%s (takes effect at next arm)",
                         core::modes::ParamNameW(p), vx, vv);
        else
            _snwprintf_s(echo, _TRUNCATE, L"%s %s=%s (takes effect at next arm)",
                         core::modes::SourceNameW(s), core::modes::ParamNameW(p),
                         (s == core::modes::Source::Vba) ? vv : vx);
        return EchoStr(echo);
    }

    // The XLOPER-grid mechanics are xlgrid.h. None of these is registered
    // thread-safe, so Excel calls them on its main thread; thread-local keeps
    // that true should one ever be.
    // 128 characters a cell: the longest answer is an error message, and one cut short misleads.
    __declspec(thread) app::XlGrid<32, 128> t_param;

    void       CellStr (int i, const wchar_t* text) { t_param.Str(i, text); }
    LPXLOPER12 CellEcho(const wchar_t* text)        { return t_param.Scalar(text); }
    LPXLOPER12 Grid    (int rows, int cols)         { return t_param.AsGrid(rows, cols); }

    LPXLOPER12 GetTraceParamBody(LPXLOPER12 srcArg, LPXLOPER12 nameArg)
    {
        // SYMMETRIC: each returns exactly the value SetTraceParam took, so a
        // cell can read it and set it back. Live capacity, drops and pauses
        // answer a different question and live on the summary.
        if (NamesWord(srcArg, nameArg, L"BUFFERSIZE"))
        {
            wchar_t sz[32]; FormatBufferW(core::modes::GetBufferBytes(), sz, 32);
            return CellEcho(sz);
        }

        if (NamesWord(srcArg, nameArg, L"BUFFERWHENFULL"))
            return CellEcho(core::modes::GetPauseOnFull() ? L"PAUSE" : L"DROP");

        if (NamesWord(srcArg, nameArg, L"FORMAT"))
            return CellEcho(core::modes::FormatNameW(core::modes::GetFormat()));

        // LOGLEVEL is source-less too, and reads back the current log level.
        if (NamesWord(srcArg, nameArg, L"LOGLEVEL"))
        {
            const char* n = core::Log::LevelName(core::Log::GetLevel());
            wchar_t w[16];
            core::WidenAscii(n, w, 16);
            return CellEcho(w);
        }

        bool both = false; core::modes::Source s = core::modes::Source::Xll;
        if (!ParseSource(srcArg, both, s))
            return CellEcho(L"#Err - Source must be XLL or VBA, or omit for both");

        const bool haveName = !IsMissing(nameArg);
        core::modes::Param p = core::modes::Param::Depth;
        if (haveName && !ParseParam(nameArg, p))
            return CellEcho(L"#Err - Name must be DEPTH, ARGS, RETVAL or OBJECTS");

        // Grid sizes derive from this list.
        const core::modes::Param all[] = { core::modes::Param::Depth, core::modes::Param::Args,
                                      core::modes::Param::RetVal, core::modes::Param::Objects };
        constexpr int kParams = static_cast<int>(sizeof(all) / sizeof(all[0]));
        // The widest grid must fit t_param, whose last cell is reserved for Scalar.
        static_assert(2 * kParams * 3 <= 31,
                      "the parameter grid no longer fits t_param; widen XlGrid");
        const core::modes::Source srcs[2] = { core::modes::Source::Xll, core::modes::Source::Vba };
        wchar_t v[24];

        // Source AND Name -> the one value.
        if (!both && haveName) { ValueTextW(s, p, v, 24); return CellEcho(v); }

        // Source only -> that source's settings: Name, Value.
        if (!both)
        {
            for (int i = 0; i < kParams; i++)
            {
                CellStr(i * 2, core::modes::ParamNameW(all[i]));
                ValueTextW(s, all[i], v, 24);
                CellStr(i * 2 + 1, v);
            }
            return Grid(kParams, 2);
        }

        // Name only -> that setting on both sources: Source, Value.
        if (haveName)
        {
            for (int i = 0; i < 2; i++)
            {
                CellStr(i * 2, core::modes::SourceNameW(srcs[i]));
                ValueTextW(srcs[i], p, v, 24);
                CellStr(i * 2 + 1, v);
            }
            return Grid(2, 2);
        }

        // Neither -> everything, growing with the settings list rather than
        // being a fixed shape.
        int c = 0;
        for (int si = 0; si < 2; si++)
            for (int pi = 0; pi < kParams; pi++)
            {
                CellStr(c++, core::modes::SourceNameW(srcs[si]));
                CellStr(c++, core::modes::ParamNameW(all[pi]));
                ValueTextW(srcs[si], all[pi], v, 24);
                CellStr(c++, v);
            }
        return Grid(2 * kParams, 3);
    }
}

extern "C" LPXLOPER12 __stdcall XRayXL_SetTraceParam(LPXLOPER12 src, LPXLOPER12 name, LPXLOPER12 value)
{
    return GuardedOper("XRayXL_SetTraceParam",
                       SetTraceParamBody, src, name, value);
}

extern "C" LPXLOPER12 __stdcall XRayXL_GetTraceParam(LPXLOPER12 src, LPXLOPER12 name)
{
    return GuardedOper("XRayXL_GetTraceParam", GetTraceParamBody, src, name);
}
