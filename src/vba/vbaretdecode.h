// A VBA FUNCTION'S RESULT, AT ANY DEPTH -- what kind it is, where it lives, and
// how to say it in one trace cell:
//
//   1. THE EXIT OPCODE NAMES THE RETURN KIND. A procedure leaves through a
//      slot chosen by its declaration, at every depth and whatever called it:
//
//         623 Byte   624 Integer/Boolean   625 Long   626 Single
//         627 Double/Date   628 Currency   630 String
//         631 Object (As Object / As Collection / As Range ...)
//         634 LongLong -- and every typed ARRAY   952 Variant
//         635 a SUB: no result exists   504 a class/form SUB
//
//      Microsoft's PDBs name these lblEX_ExitProcUI1/I2/I4/R4/R8 and
//      lblEX_ExitProc, stable across 41 VBE7 builds. Five slots share the plain
//      ExitProc handler and are still five OPCODES, so a Sub announces itself
//      and no caller gate is needed. Class and form FUNCTIONS all leave through
//      1664 and are typed from their STORE instead.
//
//   2. A SCALAR OR TYPED-ARRAY RESULT IS AT [R14-8]: scalars as raw bytes (the
//      kind says how to read them), a typed array as a live SAFEARRAY* whose
//      element VARTYPE sits four bytes before the descriptor (FADF_HAVEVARTYPE).
//      A typed String is a live BSTR there too.
//
//   3. A VARIANT RESULT IS A LIVE VARIANT AT [R14-0x18]: vt at -0x18, value at
//      -0x10, including a BSTR still allocated and a VT_ARRAY payload. The
//      caller's own result VARIANT (argument slot 1) is still EMPTY at the exit
//      opcode, so the local is the only place the answer exists.
//
//   4. AN ARRAY OF VARIANTS -- what Array(...) and Range.Value produce -- is
//      VT_ARRAY|VT_VARIANT with 24-byte elements, each a VARIANT of its own, so
//      nesting is an element whose own vt says so and the decoder recurses to a
//      bounded depth. An object renders as in the ARGUMENT column: its address or
//      `Nothing`, plus its class and detail when OBJECTS is on.
//
// Nothing here guesses; what fails validation is empty.
#pragma once
#include "vbaoleaut.h"
#include <cstdint>
#include "core/render.h"

namespace vba
{
    // The element cap is core/render.h's, shared with the XLL column.
    using core::kMaxRenderedElems;

    enum class RetKind
    {
        Unknown,      // an exit slot this table does not know: decline
        None,         // a Sub -- nothing to read, nothing to decline
        Byte, Integer, Long, Single, Double, Currency,
        String,       // a live BSTR at [R14-8]
        Object,       // an interface pointer at [R14-8], or 0 for Nothing
        LongLongOrArray, // slot 634 serves both; the store opcode splits them
        Variant,
    };

    // The declared return kind from the exit opcode the activation ended on.
    RetKind ExitReturnKind(std::uint16_t exitOp);

    // The kind implied by the instruction that STORED the result, for exits
    // carrying no type of their own (1664). Seven types are measured; anything
    // else is Unknown and counted rather than guessed.
    RetKind StoreReturnKind(std::uint16_t storeOp);
    const char* RetKindName(RetKind k);

    // One trace cell for the result of the activation whose frame base is
    // `r14`, with the kind supplied by the caller. `storeOp` is the opcode that
    // wrote [R14-8] when known, needed only to split slot 634. False, with `out`
    // empty, when nothing can be said truthfully; `typeOut` receives the name to
    // publish in `rettype`.
    bool DescribeReturnKind(std::uint64_t r14, RetKind k, std::uint16_t storeOp,
                            char* out, int cap, const char** typeOut);

    // ONE DECODER FOR BOTH COLUMNS -- the ARGUMENT side calls this too, because
    // a ByVal Variant parameter's first two frame slots are, byte for byte, a
    // VARIANT: the VARTYPE at slot k, the value at k+1.
    //
    // `heldType` is the one thing the columns need differently: `ret` has a
    // `rettype` column beside it and renders bare (`42`), `args` names the held
    // type inline (`Integer(42)`). "" when the rendering already says the type
    // or there is nothing to name.
    //
    // `maxElems` is the caller's, so a declared array and one arriving inside a
    // Variant cap alike.
    bool DescribeVariantValue(std::uint64_t at, char* out, int cap,
                              const char** heldType, int maxElems = kMaxRenderedElems);

    // THE ONE ARRAY RENDERER: "<Elem>[lo..hi,...]{e1,e2,...}" for a SAFEARRAY
    // whose header has already been validated by ReadSafeArrayHeader. Both
    // columns render through this, so an array reads the same wherever it lands.
    bool RenderSafeArrayValue(const SaInfo& sa, char* out, int cap, int maxElems);

    // ONE ELEMENT of a SAFEARRAY the CALLER has already walked and validated.
    // The two columns walk the descriptor separately -- an argument reaches an
    // array through a POINTER TO A POINTER -- but turning an element's BYTES
    // into text is shared.
    bool DescribeArrayElement(std::uint16_t vt, std::uint64_t at, char* out, int cap,
                              int maxElems = kMaxRenderedElems);

    // Latched at arm from the OBJECTS setting. Off, an object renders as its
    // address, exactly as it did before the setting existed.
    void SetDescribeObjects(bool on);

    // A BSTR AT `p`, OR NOTHING -- the ONE reader, shared by both columns.
    // `told`: something else already says this is a string (a VARIANT tag, an
    // exit opcode). Only then is a ZERO length a value.
    bool DescribeBstrValue(std::uint64_t p, char* out, int cap, bool told = false);

    // `VtName` and the VARIANT tag constants live in vbaoleaut.h, included
    // above: they are facts about the type, not about decoding a return value.
}
