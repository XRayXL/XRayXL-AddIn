#include "log.h"
#include "ascii.h"

#include <windows.h>
#include <fstream>
#include <mutex>
#include <cstring>

namespace core
{

namespace
{
    std::mutex    g_mutex;
    // Whether this thread holds g_mutex, so a contained fault can release it (ReleaseHeldByThisThread).
    __declspec(thread) int t_held = 0;
    struct Held
    {
        Held()  { g_mutex.lock(); ++t_held; }
        ~Held() { --t_held; g_mutex.unlock(); }
        Held(const Held&) = delete;
        Held& operator=(const Held&) = delete;
    };
    std::wstring  g_path;
    volatile LONG g_level = static_cast<LONG>(Log::Level::Info);   // startup default

    // Log4Net %date: comma before the millis.
    std::string Stamp()
    {
        SYSTEMTIME t{};
        GetLocalTime(&t);
        char buf[32];
        _snprintf_s(buf, _TRUNCATE, "%04d-%02d-%02d %02d:%02d:%02d,%03d",
                    t.wYear, t.wMonth, t.wDay,
                    t.wHour, t.wMinute, t.wSecond, t.wMilliseconds);
        return buf;
    }

    const char* LevelText(Log::Level l)
    {
        switch (l)
        {
        case Log::Level::Debug:   return "DEBUG";
        case Log::Level::Warning: return "WARNING";
        case Log::Level::Error:   return "ERROR";
        default:                  return "INFO";
        }
    }

    // Caller holds g_mutex. Open/write/close per line so each is flushed: a
    // crash leaves the log complete up to the last thing that happened.
    void WriteLine(const std::string& line)
    {
        std::ofstream f(g_path.c_str(), std::ios::out | std::ios::app);
        if (f.is_open()) f << line << "\n";
    }
}

namespace Log
{
    void  SetLevel(Level lvl) { InterlockedExchange(&g_level, static_cast<LONG>(lvl)); }
    Level GetLevel()          { return static_cast<Level>(InterlockedCompareExchange(&g_level, 0, 0)); }
    const char* LevelName(Level lvl) { return LevelText(lvl); }

    bool LevelFromText(const char* text, Level& out)
    {
        if (!text) return false;
        if (_stricmp(text, "DEBUG")   == 0) { out = Level::Debug;   return true; }
        if (_stricmp(text, "INFO")    == 0) { out = Level::Info;    return true; }
        if (_stricmp(text, "WARNING") == 0 ||
            _stricmp(text, "WARN")    == 0) { out = Level::Warning; return true; }
        if (_stricmp(text, "ERROR")   == 0) { out = Level::Error;   return true; }
        return false;
    }

    std::wstring Path()
    {
        const Held lock;
        return g_path;
    }

    void Open(const std::wstring& path)
    {
        const Held lock;
        // The same log again (xlAutoOpen runs again when the add-in is added again) keeps its session.
        if (!g_path.empty() && g_path == path) return;
        g_path = path;

        // Read once, before anything can call SetLevel.
        wchar_t v[16]{};
        if (GetEnvironmentVariableW(L"XRAYXL_LOGLEVEL", v, 16) > 0)
        {
            char narrow[16]{};
            core::NarrowAscii(v, narrow, sizeof(narrow));
            Level lvl;
            if (LevelFromText(narrow, lvl))
                InterlockedExchange(&g_level, static_cast<LONG>(lvl));
        }

        std::ofstream(g_path.c_str(), std::ios::out | std::ios::trunc);   // fresh per session
        const Level cur = static_cast<Level>(InterlockedCompareExchange(&g_level, 0, 0));
        char line[128];
        _snprintf_s(line, _TRUNCATE, "%s [%5lu] %-7s - XRayXL log opened (level %s)",
                    Stamp().c_str(), GetCurrentThreadId(), LevelText(Level::Info), LevelText(cur));
        WriteLine(line);
    }

    void Write(Level lvl, const std::string& msg)
    {
        if (static_cast<LONG>(lvl) < InterlockedCompareExchange(&g_level, 0, 0)) return;
        const Held lock;
        if (g_path.empty()) return;
        char head[64];
        _snprintf_s(head, _TRUNCATE, "%s [%5lu] %-7s - ",
                    Stamp().c_str(), GetCurrentThreadId(), LevelText(lvl));
        WriteLine(std::string(head) + msg);
    }

    void Debug(const std::string& m)   { Write(Level::Debug,   m); }
    void Info(const std::string& m)    { Write(Level::Info,    m); }
    void Warning(const std::string& m) { Write(Level::Warning, m); }
    void Error(const std::string& m)   { Write(Level::Error,   m); }
    void Note(const std::string& m)    { Write(Level::Info,    m); }

    void ReleaseHeldByThisThread()
    {
        if (t_held > 0) { t_held = 0; g_mutex.unlock(); }
    }

    void FaultWhileLockedForProbe()
    {
        const Held lock;
        volatile int* nowhere = nullptr;
        *nowhere = 1;
    }
}
}   // namespace core
