// Capture configuration, held per source: the XLL/VBA split, a word the user-facing parameter
// uses too.
//
//   DEPTH   OFF | TOP | ALL   how much of the call tree emits rows. TOP is the
//                             outermost call only -- the one Excel itself
//                             initiated. OFF: the source is not hooked at all.
//   ARGS    TRUE | FALSE      decode parameters
//   RETVAL  TRUE | FALSE      decode return values
//   OBJECTS TRUE | FALSE      describe a COM object argument or result --
//                             its class, and for a Range, Worksheet or
//                             Workbook the detail that identifies it
//   BREAKPOINTS TRUE | FALSE  VBA only: add a `breaks` column, how often each
//                             call stopped at a breakpoint in the editor
//
// OBJECTS is the one setting that calls the object model, on the calculating thread; OFF returns
// the tracer to pure observation.
//
// A setting changes what is emitted, never what is counted, so "did not look" and "found nothing"
// stay distinct. The arm paths read these once and latch them, so the hot path never sees a value
// change.
#pragma once
#include <cstddef>

namespace core
{

namespace modes
{
    enum class Source { Xll = 0, Vba = 1 };
    enum class Depth  { Off = 0, Top = 1, All = 2 };

    // The calling cell (xlfCaller) is not a setting: the VBA tracer always needs it to tell an error
    // that escapes into a cell from one that propagates.
    enum class Param { Depth = 0, Args = 1, RetVal = 2, Objects = 3, Breakpoints = 4 };

    Depth GetDepth (Source s);
    bool  GetArgs  (Source s);
    bool  GetRetVal(Source s);
    bool  GetObjects(Source s);
    bool  GetBreakpoints(Source s);

    void SetDepth (Source s, Depth d);
    void SetArgs  (Source s, bool on);
    void SetRetVal(Source s, bool on);
    void SetObjects(Source s, bool on);
    void SetBreakpoints(Source s, bool on);

    const wchar_t* DepthNameW(Depth d);
    const char*    DepthName (Depth d);
    const char*    SourceName(Source s);
    const wchar_t* SourceNameW(Source s);
    const char*    ParamName (Param p);
    const wchar_t* ParamNameW(Param p);
    const wchar_t* OnOffW(bool v);

    // DEPTH != OFF, for the arming paths.
    bool XllEnabled();
    bool VbaEnabled();
    // Whether a trace file opened now gets the `breaks` column: VBA traced, with BREAKPOINTS on.
    bool BreaksColumn();

    // The output buffer, a property of the file, so no Source. 0 writes synchronously; N is a byte
    // ring drained off-thread. Bytes, not MB, so a K/KB suffix can ask for a sub-MB ring.
    std::size_t GetBufferBytes();
    void        SetBufferBytes(std::size_t bytes);

    // When the ring fills: false drops the row (fast, lossy), true pauses the traced thread until a
    // slot frees (never loses, may throttle the calc). Ignored when BUFFERSIZE=0.
    bool GetPauseOnFull();
    void SetPauseOnFull(bool pause);

    // The file's format: CSV, one row a line with the values as text, or JSON Lines, one object
    // a line with every value structured and typed. A property of the file, so no Source.
    enum class Format { Csv = 0, Jsonl = 1 };
    Format GetFormat();
    void   SetFormat(Format f);
    const char*    FormatName (Format f);
    const wchar_t* FormatNameW(Format f);

    // Developer diagnostics in the disarm report (opcode and p-code dumps), off so an ordinary
    // report stays clean; XRAYXL_DIAG=1 to find why a procedure was not named or a return typed.
    bool DiagEnabled();
}
}   // namespace core
