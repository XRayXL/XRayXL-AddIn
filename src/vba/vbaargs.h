// A VBA procedure's arguments, read from its frame on 64-bit Excel.
//
//    * R14 is the frame base; the arguments are [R14, R14 + argSz), argSz the WORD at
//      trailer+0x08: argSz = 8 * (nargs + 1). Slot 0 is reserved.
//    * Each slot is eight bytes: ByVal holds the value (a BSTR pointer for String), ByRef a
//      pointer.
//    * Past argSz is the caller's frame, which reads without faulting, so the bound is ours.
//
// Types come from the bytecode (vbapcode.h), never guessed from the frame. An untyped value is
// a raw qword unless it validates itself (a BSTR, a SAFEARRAY).
#pragma once
#include <cstdint>
#include "core/textbuf.h"

namespace vba
{
    // XRAYXL_DIAG only: append `#NNN` to each named type, for when a name is present but wrong.
    void SetArgTypeOpcodeDiagnostics(bool on);

    struct ArgCapture
    {
        bool ok    = false;

        // Declared parameters, not slots: a ByVal Variant occupies three. `slots` when untyped.
        int  params = 0;

        // Argument slots, excluding slot 0. The trailer carries no parameter count, only this.
        int  slots = 0;

        // 2 for a Function returning Variant, whose caller's result VARIANT arrives first.
        int  firstSlot = 1;

        // Caller-provided (per-thread), so ArgCapture stays small on the hot-path stack. Null
        // makes CaptureArgs decline.
        core::TextBuf* text = nullptr;

        // A slot that had to be dereferenced is ByRef in substance, which the exit re-read
        // needs: a write-only `String` has no type but still points at the caller's BSTR.
        bool viaPointer = false;

        // Render only the ByRef slots: a declared `&` type, or one read through a pointer.
        bool byRefOnly = false;

        // FNV-1a over the ByRef slots as rendered, so the exit can tell whether they moved.
        std::uint64_t byRefHash = 0;
        // A ByRef slot was written as its shape alone, so the hash cannot see its contents.
        bool byRefShapeOnly = false;

        // From the p-code, e.g. "(Long,String)"; "?" where a type was not recoverable.
        char signature[256] = {};
    };

    // Counted separately: "never looked" and "the frame did not check out" are different facts.
    enum class ArgDecline
    {
        NoTrailer,
        NoFrameBase,
        ArgSzUnreadable,
        ArgSzNotMultipleOf8,
        ArgSzOutOfRange,
        SlotUnreadable,
        NoRenderBuffer,
        TooLarge,
        Count_
    };

    const char*   ArgDeclineName(ArgDecline d);

    // Opcodes that reached a parameter slot and could not be named: a hole in the type table.
    const char*   ArgTypeUnknownWarning();
    void          ResetArgCounts();

    // Never throws, never faults, never writes. False leaves `out` empty.
    bool CaptureArgs(std::uint64_t trailer, std::uint64_t r14, ArgCapture& out);

    // One line for the disarm report.
    const char* ArgsLine();
}
