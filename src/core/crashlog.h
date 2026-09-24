// Records why the process died, if it dies: an unhandled-exception filter
// writing the exception code, the faulting address and the module that owns it,
// through the raw file API alone -- no allocation, no CRT, nothing that needs a
// working heap at the point of death.
#pragma once
#include <string>
#include <cstddef>

namespace core
{

namespace crashlog
{
    void Install(const std::wstring& path);

    // First-chance capture for a jump to an address in no mapped module. Excel installs its own
    // filter over ours, but a vectored handler runs before any SEH and cannot be displaced. It
    // records the stack, since the address alone does not say who jumped there, and acts only when
    // the instruction pointer is outside every loaded module.
    void InstallVectored(const std::wstring& dumpDir);

    // The minidump is opt-in: even a lean dump of Excel carries workbook content unencrypted into
    // %TEMP%, and MiniDumpWriteDump is an EDR signature. XRAYXL_CRASHDUMP=1 is read once at
    // install; an environment variable, because the handler goes in before anything can be armed.
    // SetDumpEnabled is for the instrument that proves the capture works.
    void SetDumpEnabled(bool on);

    // Executable memory we allocated: a stub page belongs to no module, so
    // without this every stub address reads as "not mapped code".
    void NoteExecRegion(const void* base, std::size_t bytes, const char* what);

    // A plain line, so the last thing written before a crash says where we were.
    void Note(const char* text);
}
}   // namespace core
