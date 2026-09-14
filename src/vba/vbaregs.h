// THE THUNK'S SAVED-REGISTER BLOCK -- an ABI between hand-written assembly and
// C++, held together by nothing but agreement. vbathunk.asm saves the
// interpreter's registers on entry and hands the top of the block to the hook as
// `savedRegs`; the interpreter's state lies at fixed negative offsets from it.
//
// IT HAS BEEN WRONG BEFORE, AND SILENTLY: one omitted push shifts every offset
// below it by a register, and the guarded read still SUCCEEDS -- the memory is
// readable, it simply holds the wrong register. So there is exactly one copy of
// these offsets, here. If the pushes in vbathunk.asm change, change them here
// and nowhere else -- and count them.
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
