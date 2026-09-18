// The thunk's saved-register block, shared between vbathunk.asm and C++. The thunk saves the
// interpreter's registers and hands the top of the block to the hook as `savedRegs`.
//
// The only copy of these offsets. One omitted push shifts every offset below it, and the
// guarded read still succeeds on the wrong register, so if the pushes change, change them here
// and count them.
//
//   push order (after the push that establishes rbx), fifteen of them:
//     rax rcx rdx r8 r9 r10 r11 rbp rsi rdi r12 r13 r14 r15 flags
#pragma once

namespace vba
{
    // Offsets DOWN from the top of the saved block.
    constexpr int kReg_rax   = -0x08;
    constexpr int kReg_rcx   = -0x10;
    constexpr int kReg_rdx   = -0x18;
    constexpr int kReg_r8    = -0x20;
    constexpr int kReg_r9    = -0x28;
    constexpr int kReg_r10   = -0x30;
    constexpr int kReg_r11   = -0x38;
    constexpr int kReg_rbp   = -0x40;
    constexpr int kReg_rsi   = -0x48;   // the p-code instruction pointer
    constexpr int kReg_rdi   = -0x50;
    constexpr int kReg_r12   = -0x58;
    constexpr int kReg_r13   = -0x60;
    constexpr int kReg_r14   = -0x68;   // the VBA frame base
    constexpr int kReg_r15   = -0x70;
    constexpr int kReg_flags = -0x78;

    // Fifteen 8-byte saves; vbathunk.asm SAVED_BYTES (120) must match.
    constexpr int kSavedBytes = 15 * 8;
    static_assert(kReg_flags == -kSavedBytes, "vbaregs.h and vbathunk.asm SAVED_BYTES disagree");
}
