#include "tracemodes.h"
#include "notify.h"
#include <windows.h>

namespace core
{

namespace modes
{
    namespace
    {
        // Interlocked LONGs, not bools: the setters run on whatever thread
        // Application.Run arrived on and arming reads on another. Indexed by
        // Source, so a third source is a size change and nothing else.
        constexpr int kSources = 2;

        // Both sources on: a trace of a VBA-heavy workbook with no VBA rows reads as a
        // broken tool. VBA OFF stays available, since it patches the interpreter.
        volatile LONG g_depth [kSources] = { static_cast<LONG>(Depth::All),
                                             static_cast<LONG>(Depth::All) };
        // ON: a new switch must not quietly change what an existing user gets.
        volatile LONG g_args  [kSources] = { 1, 1 };
        volatile LONG g_retval[kSources] = { 1, 1 };
        // ON, and it is the only one whose cost is a CALL INTO EXCEL rather than
        // a read of memory. Off, an object renders as its address, so turning it off loses detail and changes
        // nothing else.
        volatile LONG g_objects[kSources] = { 1, 1 };
        // OFF: the column it adds changes the file's header, which every reader of a trace relies on.
        volatile LONG g_breakpoints[kSources] = { 0, 0 };

        // The shipped default is a ring -- the production path, and
        // what the suites run. 64 MB is deep enough that a functional workload
        // never laps the ~1 ms drain. LONG64, as the accessors around it are.
        volatile LONG64 g_bufferBytes = 64ll * 1024 * 1024;

        // 1 = PAUSE by default: a full ring makes the calc wait until it is half empty, and loses nothing.
        volatile LONG g_pauseOnFull = 1;

        // CSV by default: the format every reader of a trace already understands.
        volatile LONG g_format = static_cast<LONG>(Format::Csv);

        int Ix(Source s) { return (s == Source::Vba) ? 1 : 0; }

        LONG Read(volatile LONG* a, Source s)
        {
            return InterlockedCompareExchange(&a[Ix(s)], 0, 0);
        }
    }

    Depth GetDepth (Source s) { return static_cast<Depth>(Read(g_depth, s)); }
    bool  GetArgs  (Source s) { return Read(g_args,   s) != 0; }
    bool  GetRetVal(Source s) { return Read(g_retval, s) != 0; }
    bool  GetObjects(Source s) { return Read(g_objects, s) != 0; }
    bool  GetBreakpoints(Source s) { return Read(g_breakpoints, s) != 0; }

    // Every setter announces itself, so whatever displays these is told. Control path only.
    void SetDepth (Source s, Depth d) { InterlockedExchange(&g_depth[Ix(s)],  static_cast<LONG>(d)); NotifyStateChanged(); }
    void SetArgs  (Source s, bool on) { InterlockedExchange(&g_args[Ix(s)],   on ? 1 : 0); NotifyStateChanged(); }
    void SetRetVal(Source s, bool on) { InterlockedExchange(&g_retval[Ix(s)], on ? 1 : 0); NotifyStateChanged(); }
    void SetObjects(Source s, bool on) { InterlockedExchange(&g_objects[Ix(s)], on ? 1 : 0); NotifyStateChanged(); }
    void SetBreakpoints(Source s, bool on) { InterlockedExchange(&g_breakpoints[Ix(s)], on ? 1 : 0); NotifyStateChanged(); }

    const wchar_t* DepthNameW(Depth d)
    {
        switch (d) { case Depth::Top: return L"TOP"; case Depth::All: return L"ALL"; default: return L"OFF"; }
    }
    const char* DepthName(Depth d)
    {
        switch (d) { case Depth::Top: return "TOP"; case Depth::All: return "ALL"; default: return "OFF"; }
    }
    const char*    SourceName (Source s) { return (s == Source::Vba) ? "VBA" : "XLL"; }
    const wchar_t* SourceNameW(Source s) { return (s == Source::Vba) ? L"VBA" : L"XLL"; }

    const char* ParamName(Param p)
    {
        switch (p)
        {
        case Param::Args:   return "ARGS";
        case Param::RetVal: return "RETVAL";
        case Param::Objects: return "OBJECTS";
        case Param::Breakpoints: return "BREAKPOINTS";
        default:            return "DEPTH";
        }
    }
    const wchar_t* ParamNameW(Param p)
    {
        switch (p)
        {
        case Param::Args:   return L"ARGS";
        case Param::RetVal: return L"RETVAL";
        case Param::Objects: return L"OBJECTS";
        case Param::Breakpoints: return L"BREAKPOINTS";
        default:            return L"DEPTH";
        }
    }
    const wchar_t* OnOffW(bool v) { return v ? L"TRUE" : L"FALSE"; }

    bool XllEnabled() { return GetDepth(Source::Xll) != Depth::Off; }
    bool VbaEnabled() { return GetDepth(Source::Vba) != Depth::Off; }
    bool BreaksColumn() { return VbaEnabled() && GetBreakpoints(Source::Vba); }

    std::size_t GetBufferBytes() { return static_cast<std::size_t>(InterlockedCompareExchange64(&g_bufferBytes, 0, 0)); }
    void        SetBufferBytes(std::size_t bytes) { InterlockedExchange64(&g_bufferBytes, static_cast<LONG64>(bytes)); NotifyStateChanged(); }

    bool GetPauseOnFull() { return InterlockedCompareExchange(&g_pauseOnFull, 0, 0) != 0; }
    void SetPauseOnFull(bool pause) { InterlockedExchange(&g_pauseOnFull, pause ? 1 : 0); NotifyStateChanged(); }

    Format GetFormat() { return static_cast<Format>(InterlockedCompareExchange(&g_format, 0, 0)); }
    void   SetFormat(Format f) { InterlockedExchange(&g_format, static_cast<LONG>(f)); NotifyStateChanged(); }
    const char*    FormatName (Format f) { return f == Format::Jsonl ? "JSONL" : "CSV"; }
    const wchar_t* FormatNameW(Format f) { return f == Format::Jsonl ? L"JSONL" : L"CSV"; }

    // Cached for the life of the process: a ship-vs-investigate choice, not
    // something to flip mid-session.
    bool DiagEnabled()
    {
        static const bool on = []
        {
            wchar_t v[8]{};
            return GetEnvironmentVariableW(L"XRAYXL_DIAG", v, 8) > 0 && v[0] == L'1';
        }();
        return on;
    }
}
}   // namespace core
