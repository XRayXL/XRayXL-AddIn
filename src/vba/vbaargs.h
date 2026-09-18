// Reading a VBA procedure's arguments out of its frame, on 64-bit Excel.
//
//    * R14 is the frame base. The argument region is [R14, R14 + argSz), and argSz is the
//      WORD at trailer+0x08, in bytes: argSz = 8 * (nargs + 1).
//    * Slot 0, at [R14+0], is reserved. Arguments are slots 1..n.
//    * Every slot is eight bytes. ByVal Long and Integer hold the value, ByVal Double the
//      raw IEEE bits, ByVal String a BSTR pointer; ByRef of any type holds a pointer.
//    * One slot past argSz is the caller's frame. Reading it does not fault, so the bound
//      is enforced here.
//
// The type is not in the frame and is never guessed from it: it comes from the procedure's own
// bytecode (vbapcode.h). A parameter the body never reads has no recoverable type. It renders
// "?" and its value is a raw qword, or is decoded only where it validates itself (a BSTR, a
// SAFEARRAY).
#pragma once
#include <cstdint>

namespace vba
{
    // XRAYXL_DIAG only: append `#NNN` to each named type, naming the opcode that
    // named it. `?opNNN` says which opcode failed; this says which succeeded,
    // for when a name is present but wrong.
    void SetArgTypeOpcodeDiagnostics(bool on);

    struct ArgCapture
    {
        bool ok    = false;

        // DECLARED PARAMETERS -- what a reader means by "how many arguments". Not
        // the same as `slots`: a ByVal Variant occupies three. Falls back to `slots`
        // when no types were recovered.
        int  params = 0;

        // ARGUMENT SLOTS, excluding the reserved slot 0. Almost every parameter
        // is one 8-byte slot, but a `ByVal Variant` is a 24-byte VARIANT and
        // occupies THREE. Nothing in the trailer carries the declared parameter
        // count, so this reports what is actually known.
        int  slots = 0;

        // Which slot the first argument is at. Normally 1, because slot 0 is reserved; 2 for a
        // Function returning Variant, where the caller's result VARIANT arrives first. Labels
        // count arguments (a1, a2, ...) while reads are done at slots.
        int  firstSlot = 1;

        // CALLER-PROVIDED (a large per-thread render buffer), so a big
        // array or variant argument is not trimmed to a fixed inline size and
        // ArgCapture stays small on the hot-path stack. The render stops at
        // `textCap`; a null `text` makes CaptureArgs decline cleanly.
        char* text    = nullptr;
        int   textCap = 0;

        // A slot that had to be DEREFERENCED is ByRef in substance, whatever the
        // p-code declared -- the property the exit re-read needs. A
        // write-only `String` recovers no type at all, but its slot still points
        // at the caller's BSTR.
        bool viaPointer = false;

        // Render only the ByRef slots: a declared `&` type, or one read through a pointer.
        bool byRefOnly = false;

        // FNV-1a over the ByRef slots as rendered, so the exit can tell whether they moved.
        std::uint64_t byRefHash = 0;

        // Recovered from the procedure's own p-code, e.g. "(Long,String)". A
        // position reads "?" when that parameter's type was not recoverable --
        // almost always because the body never reads it.
        char signature[256] = {};
    };

    // Counted, never merged: "we never looked" and "we looked and the frame did
    // not check out" are different facts about the tracer.
    enum class ArgDecline
    {
        NoTrailer,
        NoFrameBase,
        ArgSzUnreadable,
        ArgSzNotMultipleOf8,
        ArgSzOutOfRange,
        SlotUnreadable,
        NoRenderBuffer,
        Count_
    };

    const char*   ArgDeclineName(ArgDecline d);

    // Opcodes that reached a parameter slot and could not be NAMED -- a hole in
    // the type table, not an absence of evidence. Distinct from a bare `?`,
    // which means no instruction referenced the slot at all.
    const char*   ArgTypeUnknownWarning();
    void          ResetArgCounts();

    // Never throws, never faults, never writes. False leaves `out` empty.
    bool CaptureArgs(std::uint64_t trailer, std::uint64_t r14, ArgCapture& out);

    // One line for the disarm report.
    const char* ArgsLine();
}
