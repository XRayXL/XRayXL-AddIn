#pragma once
#include <cstdint>

// THE VBE7 PROCEDURE TRAILER: a small structure just past the p-code body that
// describes the procedure, and the single authority for two kinds of offset --
// WHERE the pointer sits relative to the interpreter's dispatch stack pointer,
// and which FIELDS sit where inside it. (vbaregs.h is the sister ABI, and says
// what a scattered layout costs.)
//
// Every read is GUARDED and every failure COUNTED: these offsets are the magic
// numbers on this path, and a field's WIDTH matters as much as its offset.
namespace vba
{
    // WHERE THE TRAILER POINTER LIVES: [dispatchSp + kTrailerOffset]. Published
    // prior art agrees; the one magic number on the walk.
    constexpr std::uint32_t kTrailerOffset = 0xB8;

    // trailer+0x08, a WORD read in BYTES: 8/16/24/32 for 0/1/2/3 arguments. Read
    // as a DWORD it picks up the neighbouring word.
    constexpr std::uint32_t kTrl_argSz    = 0x08;

    // trailer+0x0C is ProcSize, the code length in BYTES, so the bytecode is
    // [trailer - ProcSize, trailer). Measured 32/32/40/48 for small bodies.
    // Also a WORD; every reader must use that width.
    constexpr std::uint32_t kTrl_procSize = 0x0C;

    // The largest value the WORD can hold.
    constexpr std::uint32_t kTrl_procSizeMax = 0xFFFF;
}
