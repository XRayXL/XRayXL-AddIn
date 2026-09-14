#pragma once
#include <windows.h>

// The register snapshot xllthunk.asm hands to the C++ recorders.
// Layout is shared with xllthunk.asm; the static_asserts keep them in step.
// The thunk keeps its argument count in the 8 bytes after this block.

namespace xll
{
    struct Regs
    {
        ULONG64  ireg[4];      // +0x00  rcx, rdx, r8, r9  -- integer/pointer args 0..3
        double   xmm[4];       // +0x20  xmm0..xmm3        -- floating args, BY POSITION
        ULONG64  rax;          // +0x40  integer return
        double   xmmRet;       // +0x48  xmm0 on return
        ULONG64* stackArgs;    // +0x50  -> the caller's 5th argument onward
        ULONG64  spare;        // +0x58  pad to 0x60
    };

    static_assert(sizeof(Regs) == 0x60, "xllthunk.asm assumes sizeof(Regs) == 0x60");
    static_assert(offsetof(Regs, ireg)      == 0x00, "xllthunk.asm assumes ireg at +0x00");
    static_assert(offsetof(Regs, xmm)       == 0x20, "xllthunk.asm assumes xmm at +0x20");
    static_assert(offsetof(Regs, rax)       == 0x40, "xllthunk.asm assumes rax at +0x40");
    static_assert(offsetof(Regs, xmmRet)    == 0x48, "xllthunk.asm assumes xmmRet at +0x48");
    static_assert(offsetof(Regs, stackArgs) == 0x50, "xllthunk.asm assumes stackArgs at +0x50");

    // Position picks the index, type picks the register file -- they are not
    // competing numbering schemes: the SECOND argument of f(int, double, ...)
    // is in XMM1, not XMM0. Beyond the fourth, arguments are on the stack.
    inline ULONG64 IntArgAt(const Regs& r, int pos)
    {
        if (pos < 4) return r.ireg[pos];
        return r.stackArgs[pos - 4];
    }
    inline double DoubleArgAt(const Regs& r, int pos)
    {
        if (pos < 4) return r.xmm[pos];
        ULONG64 raw = r.stackArgs[pos - 4];
        double d; memcpy(&d, &raw, sizeof(d));
        return d;
    }
}
