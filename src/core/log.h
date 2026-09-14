// The process log -- %TEMP%\XRayXL\Logs\XRayXL_<pid>.log, one file per process,
// written in order so it reads as a timeline rather than a dump. Levelled, in
// Log4Net line format:
//   2026-09-07 10:53:07,123 [ 4812] INFO    - armed 6 of 6 registered
//
// The namespace is capitalised because a lowercase `log` collides with ::log,
// which <cmath> puts in the global namespace (MSVC C2757).
//
// Off the hot path: control threads only, under a mutex. Fault forensics is
// crashlog's job and obeys different rules (docs/Implementation.md).
#pragma once
#include <string>

namespace core
{

namespace Log
{
    enum class Level { Debug = 0, Info = 1, Warning = 2, Error = 3 };

    // Truncates and opens the log for this session. XRAYXL_LOGLEVEL is read
    // ONCE, here, so start-up can be followed at DEBUG before any SetLevel call
    // has had the chance to run.
    void Open(const std::wstring& path);

    // Settable at any time, including while armed: it controls the LOG, not the
    // trace (the LOGLEVEL trace parameter).
    void  SetLevel(Level lvl);
    Level GetLevel();
    const char* LevelName(Level lvl);                  // "DEBUG".."ERROR"
    bool  LevelFromText(const char* text, Level& out); // case-insensitive parse

    void Write(Level lvl, const std::string& msg);
    void Debug(const std::string& msg);
    void Info(const std::string& msg);
    void Warning(const std::string& msg);
    void Error(const std::string& msg);
    void Note(const std::string& msg);                 // == Info

    // XRAYXL_DIAG instrument: faults while holding the log's lock (XRayXL_FaultProbe "LOGLOCK").
    void FaultWhileLockedForProbe();

    // After a contained fault: releases the log's lock if this thread held it.
    void ReleaseHeldByThisThread();
}
}   // namespace core
