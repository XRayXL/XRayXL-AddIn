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

    // The trailer's +0 is its module's parent structure, which holds the module's constant pool at
    // +0x60: the pointer a running frame holds at [rbp-0xA0]. A call's pool entry names its callee
    // at +8.
    constexpr std::uint32_t kPar_pool        = 0x60;
    constexpr std::uint32_t kPoolEntryCallee = 0x08;

}
