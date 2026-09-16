// Reading a VBA procedure's arguments out of its frame.
//
// Measured on 64-bit Excel with planted sentinels, over thirteen procedures of
// known signature [measured: frame-layout]:
//
//   * R14 is the frame base. The argument region is [R14, R14 + argSz), and
//     argSz is the WORD at trailer+0x08, in BYTES. Across procedures taking
//     0/1/2/3 arguments it read 8/16/24/32 -- so argSz = 8 * (nargs + 1), an
//     AFFINE relation, not the plain multiple the prior art suggested.
//   * Slot 0, at [R14+0], is reserved: it held the same pointer in every
//     record, including procedures with no arguments. Arguments are slots
//     1..n, at [R14+8], [R14+16], ...
//   * Every slot is eight bytes. ByVal Long and Integer hold the VALUE; ByVal
//     Double holds the raw IEEE bits; ByVal String holds a BSTR pointer;
//     ByRef of any type holds a pointer to the value.
//   * argSz bounds the region EXACTLY. One slot past it is the CALLER's frame:
//     a two-argument procedure showed, at [R14+24], the variable its caller had
//     assigned immediately before the call. Reading past argSz does not fault --
//     it silently reports the caller's data as this call's argument, which is
//     why the bound is enforced here rather than trusted.
//
// THE TYPE IS NOT IN THE FRAME AND IS NOT GUESSED FROM IT: eight bytes reading
// 0x40A5790000000000 are a Double, a very large Long, or a pointer. The type
// comes from the procedure's own bytecode (vbapcode.h) and never from the shape of
// the value. Two consequences:
//   * A parameter the body never READS emits no typed load and so has no
//     recoverable type. Its position renders "?" and its value a raw qword.
//   * Where no type is known the value is still emitted -- as that raw qword, or
//     decoded only where it validates ITSELF (a BSTR carries its own length
//     prefix, a SAFEARRAY its dimensions and element type). A confident wrong
//     argument is worse than a hex number, which is obviously undecoded.
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

        // WHICH SLOT THE FIRST ARGUMENT IS AT. Normally 1, because slot 0 is
        // reserved. TWO for a Function returning `Variant`: the caller passes
        // the address of its result VARIANT, which the calling convention
        // pushes last and so arrives first, and it is not an argument. The
        // rendered labels count arguments (a1, a2, ...) while the reads are
        // done at slots, so the two must not be conflated.
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
