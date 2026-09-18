#pragma once
#include <cstdint>

// The VBE7 procedure trailer: a small structure just past the p-code body. The single authority
// for where its pointer sits relative to the dispatch stack pointer and which fields sit where
// inside it. Every read is guarded and every failure counted; a field's width matters as much
// as its offset.
namespace vba
{
    // WHERE THE TRAILER POINTER LIVES: [dispatchSp + kTrailerOffset]. Published
    // prior art agrees; the one magic number on the walk.
    constexpr std::uint32_t kTrailerOffset = 0xB8;

    // trailer+0x08, a WORD read in BYTES: 8/16/24/32 for 0/1/2/3 arguments. Read
    // as a DWORD it picks up the neighbouring word.
    constexpr std::uint32_t kTrl_argSz    = 0x08;

    // trailer+0x0C is ProcSize, the code length in bytes, so the bytecode is [trailer -
    // ProcSize, trailer). A WORD; every reader must use that width.
    constexpr std::uint32_t kTrl_procSize = 0x0C;

    // The largest value the WORD can hold.
    constexpr std::uint32_t kTrl_procSizeMax = 0xFFFF;
}
