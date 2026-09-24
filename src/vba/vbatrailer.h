#pragma once
#include <cstdint>

// The VBE7 procedure trailer, just past the p-code body: the one copy of its offsets. A field's
// width matters as much as its offset.
namespace vba
{
    // The trailer pointer is at [dispatchSp + kTrailerOffset].
    constexpr std::uint32_t kTrailerOffset = 0xB8;

    // A WORD in bytes: 8/16/24/32 for 0/1/2/3 arguments. Read as a DWORD it picks up its neighbour.
    constexpr std::uint32_t kTrl_argSz    = 0x08;

    // ProcSize: the bytecode is [trailer - ProcSize, trailer). A WORD; every reader must use that width.
    constexpr std::uint32_t kTrl_procSize = 0x0C;

    constexpr std::uint32_t kTrl_procSizeMax = 0xFFFF;

}
