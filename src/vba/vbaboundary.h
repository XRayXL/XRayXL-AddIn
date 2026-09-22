// Where a VBA activation begins and ends. vbatrace.cpp asks these two questions and manages the
// shadow stack without knowing how they are answered.
//
// Start is the statement whose instruction is the first byte of the procedure: a prologue,
// which no loop or label branches back to. Opening on the stack pointer fails, because GoSub
// moves rsp within one activation.
//
// End is an exit opcode, except the GoSub `Return` (620) and the two heap-type cleanup exits
// (1662, 1663), which fire before the real end at 1664. A deny-list, because an allow-list
// would stop closing for a procedure kind nobody tested. `End` fires no exit opcode at all and
// is handled by hooking its own opcode (slot 619).
//
// A slot index means the same opcode on every build. That 620 is the GoSub `Return` and
// 1662/1663 are cleanup is known for 7.1.11.58 only.
//
// Both functions are pure reads of the interpreter's registers and bytecode; the caller counts
// the outcomes.
#pragma once
#include <cstdint>

namespace vba
{
    // `Unavailable` is a THIRD answer, never folded into `No`: RSI or ProcSize
    // could not be read, so the question was not answered rather than answered
    // negatively. The caller degrades and says so.
    enum class Activation { No, Yes, Unavailable };
    Activation ActivationStart(std::uint64_t trailer, std::uint64_t savedRegs);

    // Does this exit dispatch end the procedure? No is a GoSub `Return`. Unreadable is treated
    // as No by the caller: a missed close is corrected by the next prologue, the stack pointer
    // or the flush, whereas a wrong close writes a row for a call VBA never made.
    enum class Ending { No, Yes, Unreadable };
    // `opOut` receives the exit opcode. The exits are TYPED -- a procedure
    // leaves through a slot chosen by its declared return kind (vbaretdecode.h) --
    // so the number names the return type at the one moment the result can be
    // read.
    Ending ProcedureEnd(std::uint64_t savedRegs, std::uint16_t* opOut = nullptr);

    // The 4-byte operand of the exit being dispatched. Only for an exit known to carry one.
    bool ExitOperand(std::uint64_t savedRegs, std::int32_t& out);
}
