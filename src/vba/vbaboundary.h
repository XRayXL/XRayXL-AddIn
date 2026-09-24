// Where a VBA activation begins and ends.
//
// Start is the statement at the procedure's first byte, a prologue no loop or label branches
// back to; the stack pointer cannot open frames, because GoSub moves rsp within one activation.
//
// End is an exit opcode, except the GoSub `Return` (620) and the two heap-type cleanup exits
// (1662, 1663), which fire before the real end at 1664. A deny-list, because an allow-list
// would stop closing for a procedure kind nobody tested. `End` fires no exit opcode at all and
// is handled by hooking its own opcode (slot 619). The roles of 620 and 1662/1663 are known
// for VBE7 7.1.11.58 only.
#pragma once
#include <cstdint>

namespace vba
{
    // `Unavailable` is never folded into `No`: the question was not answered, and the caller
    // degrades and says so.
    enum class Activation { No, Yes, Unavailable };
    Activation ActivationStart(std::uint64_t trailer, std::uint64_t savedRegs);

    // Does this exit dispatch end the procedure? No is a GoSub `Return`. Unreadable is treated
    // as No by the caller: a missed close is corrected by the next prologue, the stack pointer
    // or the flush, whereas a wrong close writes a row for a call VBA never made.
    enum class Ending { No, Yes, Unreadable };
    // `opOut` receives the exit opcode, which names the return type: a procedure leaves through
    // a slot chosen by its declared return kind (vbaretdecode.h).

    Ending ProcedureEnd(std::uint64_t savedRegs, std::uint16_t* opOut = nullptr);

    // The 4-byte operand of the exit being dispatched. Only for an exit known to carry one.
    bool ExitOperand(std::uint64_t savedRegs, std::int32_t& out);
}
