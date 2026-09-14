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

    // FIRST-CHANCE capture for the one crash the filter above cannot describe:
    // execution transferred to an address in NO mapped module. Excel installs
    // its own filter over ours, and a vectored handler runs before any SEH and
    // cannot be displaced, so this is the only place that sees it. It records
    // the faulting STACK, because an address alone says nothing about who
    // jumped there.
    //
    // Deliberately narrow -- it acts only when the instruction pointer is
    // outside every loaded module, so the reads that fault by design inside our
    // own hooks cost nothing. Always EXCEPTION_CONTINUE_SEARCH.
    void InstallVectored(const std::wstring& dumpDir);

    // THE MINIDUMP IS OPT-IN AND OFF BY DEFAULT: even the lean form
    // pulls in memory reachable from the stack, so a dump of Excel carries
    // workbook content -- positions, client names, P&L -- unencrypted into
    // %TEMP%, and MiniDumpWriteDump is an EDR signature in its own right. The
    // text capture above carries none of that.
    //
    // XRAYXL_CRASHDUMP=1, read once at install -- an environment variable
    // rather than a trace parameter because the handler goes in at load, long
    // before anything can be armed. SetDumpEnabled is for the instrument that
    // proves the capture still works.
    void SetDumpEnabled(bool on);

    // Executable memory WE allocated: a stub page belongs to no module, so
    // without this every stub address reads as "not mapped code".
    void NoteExecRegion(const void* base, std::size_t bytes, const char* what);

    // A plain line, so the last thing written before a crash says where we
    // were.
    void Note(const char* text);
}
}   // namespace core
