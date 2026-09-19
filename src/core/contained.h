// A fault caught at one of our entry points, said where it can be read. The filter half runs
// inside __except(...) and formats the line while the exception record exists; the handler half
// writes it to the log at ERROR and to the crash notes.
#pragma once
#include <windows.h>

namespace core
{
namespace contained
{
    // As the __except filter: "<who> FAULTED in <where>: code 0x... at <addr> (<module>+0x...)
    // -- <after>". Always EXCEPTION_EXECUTE_HANDLER.
    int Note(EXCEPTION_POINTERS* xp, const char* who, const char* where, const char* after);

    // In the __except block, once the thread's locks are released: logs the line Note made.
    void Report();
}
}   // namespace core
