#pragma once
#include <cstdint>

// Dispatch-table slots named in more than one file. A slot index means the same
// opcode on every VBE7 build; see vbaderive.h.
namespace vba
{
    // BosStub: the whole 16-byte body of an uncompiled procedure. VBA compiles before running,
    // so a walk never meets one; the DIAG corpus records it as a `stub` line.
    constexpr std::uint32_t kSlot_BosStub                = 0x1348 / 8;   // 617
    constexpr std::uint32_t kBosStubBody                 = 16;
    // `Stop` pauses in the editor and resumes, so it ends no frame; its BoS has already run.

    constexpr std::uint32_t kSlot_Stop                   = 0x1328 / 8;   // 613
    // GoSub `Return`: jumps back to its GoSub, in the middle of a procedure.
    constexpr std::uint32_t kSlot_GoSubReturn            = 0x1360 / 8;   // 620
    // ZeroRetVal and ZeroRetValVar (their PDB names): clear a String/Object or a
    // Variant temporary before the real end, so neither ends a procedure.
    constexpr std::uint32_t kSlot_ZeroRetVal             = 0x33F0 / 8;   // 1662
    constexpr std::uint32_t kSlot_ZeroRetValVar          = 0x33F8 / 8;   // 1663
    // `Exit For` out of a `For Each` (ExitForCollAd x3, ExitForAryVar, ExitForVar): it releases
    // the enumerator and leaves the loop, so it ends no procedure and names no return type.
    constexpr std::uint32_t kSlot_ExitForFirst          = 0x2070 / 8;   // 1038
    constexpr std::uint32_t kSlot_ExitForLast           = 0x2090 / 8;   // 1042
    inline bool IsLoopExitSlot(std::uint32_t slot) { return slot >= kSlot_ExitForFirst && slot <= kSlot_ExitForLast; }
    // Class and form Function exits, whose operand is the result slot.
    constexpr std::uint32_t kSlot_ExitProcCbHresult      = 0x3400 / 8;   // 1664
    constexpr std::uint32_t kSlot_ExitProcFrameCbHresult = 0x3408 / 8;   // 1665
}
