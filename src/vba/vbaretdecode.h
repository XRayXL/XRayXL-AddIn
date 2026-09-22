// A VBA Function's result: what kind it is, where it lives, and how to say it in one trace
// cell.
//
//    1. The exit opcode names the return kind:
//
//          623 Byte   624 Integer/Boolean   625 Long   626 Single
//          627 Double/Date   628 Currency   630 String
//          631 Object   634 LongLong and every typed array   952 Variant
//          635 a Sub   504 a class/form Sub
//
//       Class and form Functions all leave through 1664 and are typed from their store.
//
//    2. A scalar or typed-array result is at [R14-8]: scalars as raw bytes, a typed array
//       as a live SAFEARRAY*, a String as a live BSTR.
//
//    3. A Variant result is a live VARIANT at [R14-0x18]: vt at -0x18, value at -0x10. The
//       caller's own result VARIANT (argument slot 1) is still empty at the exit opcode.
//
//    4. An array of Variants is VT_ARRAY|VT_VARIANT with 24-byte elements, so the decoder
//       recurses to a bounded depth. An object renders as in the argument column.
//
// Nothing here guesses; what fails validation writes nothing and returns false. Values go to a
// core::ValueWriter, which owns the spelling.
#pragma once
#include "vbaoleaut.h"
#include <cstdint>
#include "core/valuewriter.h"

namespace vba
{
    enum class RetKind
    {
        Unknown,      // an exit slot this table does not know: decline
        None,         // a Sub -- nothing to read, nothing to decline
        Byte, Integer, Long, Single, Double, Currency,
        String,       // a live BSTR at [R14-8]
        Object,       // an interface pointer at [R14-8], or 0 for Nothing
        LongLongOrArray, // slot 634 serves both; the store opcode splits them
        Variant,
        Record,       // a user-defined Type: built in the frame, copied to the caller's buffer,
                      // which arrives as the hidden first argument at [R14+8]
        RecordInFrame,  // a Type of 1, 2, 4 or 8 bytes: returned in RAX, no hidden argument;
                        // the record is in the frame at R14 + the exit's operand
    };

    // The declared return kind from the exit opcode the activation ended on.
    RetKind ExitReturnKind(std::uint16_t exitOp);

    // The kind implied by the instruction that stored the result, for exits carrying no type of
    // their own (1664). Anything unrecognised is Unknown and counted.
    RetKind StoreReturnKind(std::uint16_t storeOp);
    const char* RetKindName(RetKind k);

    // The result of the activation whose frame base is `r14`, with the kind supplied by the
    // caller. `storeOp` is the opcode that wrote [R14-8] when known, needed only to split slot
    // 634. `exitOperand` is the exit instruction's operand, needed only by RecordInFrame. False,
    // having written nothing, when nothing can be said truthfully; `typeOut` receives the name
    // to publish in `rettype`.
    bool DescribeReturnKind(std::uint64_t r14, RetKind k, std::uint16_t storeOp,
                            std::int32_t exitOperand, core::ValueWriter& w, const char** typeOut);

    // One decoder for both columns: a ByVal Variant parameter's first two frame slots are, byte
    // for byte, a VARIANT. The held value is written inside BeginVariant/EndVariant, so a
    // number other than Double is named wherever the Variant lands.
    bool DescribeVariantValue(std::uint64_t at, core::ValueWriter& w);

    // The one array renderer, for a SAFEARRAY whose header ReadSafeArrayHeader has already
    // validated. Both columns render through this, so an array reads the same wherever it lands.
    bool RenderSafeArrayValue(const SaInfo& sa, core::ValueWriter& w);

    // One value of VARTYPE `vt` at `at`: an array element, or a declared scalar argument.
    bool DescribeArrayElement(std::uint16_t vt, std::uint64_t at, core::ValueWriter& w);

    // Latched at arm from the OBJECTS setting. Off, an object renders as its
    // address, exactly as it did before the setting existed.
    void SetDescribeObjects(bool on);

    // A BSTR AT `p`, OR NOTHING -- the ONE reader, shared by both columns.
    // `told`: something else already says this is a string (a VARIANT tag, an
    // exit opcode). Only then is a ZERO length a value.
    bool DescribeBstrValue(std::uint64_t p, core::ValueWriter& w, bool told = false);

    // `VtName` and the VARIANT tag constants live in vbaoleaut.h, included
    // above: they are facts about the type, not about decoding a return value.
}
