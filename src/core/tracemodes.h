// WHAT GETS TRACED -- capture configuration, held per SOURCE (this codebase's
// word for the XLL/VBA split, so the user-facing parameter uses it too).
//
//   DEPTH   OFF | TOP | ALL   how much of the call tree EMITS rows. TOP is the
//                             outermost call only -- the one Excel itself
//                             initiated. OFF: the source is not hooked at all.
//   ARGS    TRUE | FALSE      decode parameters
//   RETVAL  TRUE | FALSE      decode return values
//   OBJECTS TRUE | FALSE      describe a COM object argument or result --
//                             its class, and for a Range, Worksheet or
//                             Workbook the detail that identifies it
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

    // Named Param, not Setting, so it reads as the function that sets it does.
    // The calling cell (xlfCaller) is always resolved: the VBA tracer needs it to
    // tell an error that escapes into a cell from one that propagates, so it is
    // not a setting.
    enum class Param { Depth = 0, Args = 1, RetVal = 2, Objects = 3 };

    Depth GetDepth (Source s);
    bool  GetArgs  (Source s);
    bool  GetRetVal(Source s);
    bool  GetObjects(Source s);

    void SetDepth (Source s, Depth d);
    void SetArgs  (Source s, bool on);
    void SetRetVal(Source s, bool on);
    void SetObjects(Source s, bool on);

    const wchar_t* DepthNameW(Depth d);
    const char*    DepthName (Depth d);
    const char*    SourceName(Source s);
    const wchar_t* SourceNameW(Source s);
    const char*    ParamName (Param p);
    const wchar_t* ParamNameW(Param p);
    const wchar_t* OnOffW(bool v);

    // DEPTH != OFF, for the arming paths. One vocabulary for both sources, not
    // two.
    bool XllEnabled();
    bool VbaEnabled();

    // OUTPUT BUFFER, in BYTES -- a property of the FILE, not of a source, so no
    // Source. 0 is the synchronous write; N is a byte ring of that budget
    // drained off-thread. Bytes rather than MB so the command surface can
    // accept a K/KB suffix for a sub-MB ring.
    std::size_t GetBufferBytes();
    void        SetBufferBytes(std::size_t bytes);

    // WHEN THE RING FILLS: false = DROP the row (fast, lossy), true = PAUSE the
    // traced thread until a slot frees (never loses, may throttle the calc).
    // Ignored when BUFFERSIZE=0.
    bool GetPauseOnFull();
    void SetPauseOnFull(bool pause);

    // Developer diagnostics in the disarm report -- raw opcode dumps, the
    // identity struct-walk, the untyped-procedure p-code dump. Off so a
    // ordinary report stays clean; XRAYXL_DIAG=1 when investigating why a
    // procedure could not be named or a return typed.
    bool DiagEnabled();
}
}   // namespace core
