#pragma once
#include <windows.h>
#include <string>

// Parses an XLL registration type string into a decode plan. Two rules:
//
//    1. The Excel-2007 codes are two characters: C% D% F% G% K% O%.
//    2. Type O / O% is one type code occupying three ABI slots (rows, columns, array),
//       so the type index and the ABI slot index are different numbers.
//
// double f(string, double[], double) registers as "QD%K%E": three arguments, not five.

namespace xll
{
    // Named for the decode action rather than the type letter, because several
    // letters share one action.
    enum class Kind : unsigned char
    {
        Unknown = 0,
        Double,         // B          -- the value is in XMM[position]
        Int16,          // A I        -- low 16 bits of the integer register, signed
        UInt16,         // H          -- low 16 bits, unsigned
        Int32,          // J          -- low 32 bits
        PtrInt16,       // L M        -- pointer to short (L is a boolean)
        PtrInt32,       // N          -- pointer to int32
        PtrDouble,      // E          -- pointer to double
        StrAsciiZ,      // C F        -- pointer to null-terminated ASCII
        StrAsciiCount,  // D G        -- pointer to byte-counted ASCII
        StrWideZ,       // C% F%      -- pointer to null-terminated UTF-16
        StrWideCount,   // D% G%      -- pointer to WCHAR-counted UTF-16
        Fp,             // K          -- FP    {u16 rows, u16 cols, double[]}
        Fp12,           // K%         -- FP12  {i32 rows, i32 cols, double[]}
        OperNarrow,     // P R        -- XLOPER   (24 bytes, xltype at +16)
        OperWide,       // Q U        -- XLOPER12 (32 bytes, xltype at +24)
        ArrayTriple,    // O O%       -- three slots: rows*, cols*, double[]
        AsyncHandle,    // X          -- async handle; do not time this call
        Void            // no return value
    };

    struct Slot
    {
        Kind kind = Kind::Unknown;
        // 0-based ABI position, which decides the register: ireg[pos] for an
        // integer/pointer, xmm[pos] for a Double.
        int  abiIndex = 0;
        // The literal code, for the record and for diagnosis ("K%", "O", "B").
        char code[4] = { 0, 0, 0, 0 };
        // ArrayTriple only: this slot is the first of three.
        bool tripleHead = false;
    };

    struct Plan
    {
        // slotCount is what the thunk forwards and must be exact; describedCount is
        // what the args column names. Describing every slot would bloat every Target.
        static const int kMaxAbiSlots  = 768;   // 255 args x 3 slots for O, rounded up
        static const int kMaxDescribed = 64;

        Slot slots[kMaxDescribed];
        int  describedCount = 0;     // slots carrying a Slot entry: min(slotCount, kMaxDescribed)
        int  slotCount = 0;          // ABI slots, not type codes

        Kind returnKind = Kind::Unknown;
        char returnCode[4] = { 0, 0, 0, 0 };

        // The typetext and rettype columns, built here so the registration is parsed once:
        // the argument codes joined by commas ("B,B"), and the return code with its flags ("Q$").
        char signature[192] = {};
        char returnText[8] = {};
        int  paramCount = 0;         // type codes, so an O counts once though it takes three slots

        // Modifiers and modify-in-place returns affect registration, not decoding.
        // async comes from an X argument; '>' alone can be a legacy synchronous void.
        bool async = false;
        bool ok = false;             // parsed cleanly end to end
        char firstBadCode = 0;       // the code that stopped us, if !ok
    };

    // "QD%K%E", "BBB", "RPPP$", ">QX". An unrecognised code leaves ok false and firstBadCode set,
    // and the caller must decline rather than trace a function it cannot read.
    Plan Parse(const wchar_t* typeText);

}
