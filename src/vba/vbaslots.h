#pragma once
#include <cstdint>

// Dispatch-table slots named in more than one file. A slot index means the same
// opcode on every VBE7 build; see vbaderive.h.
namespace vba
{
    // BosStub: an UNCOMPILED procedure, whatever it holds. Its whole body is one of these and 16
    // bytes: the opcode, a zero dword where Bos carries its next-statement offset, a zero word, a
    // dword that rises by 0x58 per procedure through the module, and a zero dword. VBA compiles a
    // procedure before running it, so a walk never meets one; the DIAG corpus keeps such a body as
    // a `stub` line, the only place this form is measured.
    constexpr std::uint32_t kSlot_BosStub                = 0x1348 / 8;   // 617
    constexpr std::uint32_t kBosStubBody                 = 16;           // measured, every one
    // `Stop`: a breakpoint written into the source. It pauses in the editor and resumes, so it
    // ends no frame, and its statement's BoS has already run by the time it dispatches.
    constexpr std::uint32_t kSlot_Stop                   = 0x1328 / 8;   // 613
    // GoSub `Return`: jumps back to its GoSub, in the middle of a procedure.
    constexpr std::uint32_t kSlot_GoSubReturn            = 0x1360 / 8;   // 620
    // ZeroRetVal and ZeroRetValVar (their PDB names): clear a String/Object or a
    // Variant temporary before the real end, so neither ends a procedure.
    constexpr std::uint32_t kSlot_ZeroRetVal             = 0x33F0 / 8;   // 1662
    constexpr std::uint32_t kSlot_ZeroRetValVar          = 0x33F8 / 8;   // 1663
    // Class and form Function exits, whose operand is the result slot.
    constexpr std::uint32_t kSlot_ExitProcCbHresult      = 0x3400 / 8;   // 1664
    constexpr std::uint32_t kSlot_ExitProcFrameCbHresult = 0x3408 / 8;   // 1665
}
