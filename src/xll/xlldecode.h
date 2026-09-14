#pragma once
#include "xlltypeplan.h"
#include "xllregs.h"

// Rendering a captured argument or return value as short readable text.
//
// Every pointer here came from somebody else's add-in and is dereferenced under
// SEH. A value that cannot be read confidently is written EMPTY, never guessed:
// a range check that passes on nonsense yields a readable, complete-looking,
// entirely fabricated string.

namespace xll
{
    // Longest text we will write for one value. Truncation is marked with a
    // trailing '~' so a clipped value cannot be mistaken for a short one.
    const int kValueMax = 128;

    // Room to render kValueMax characters escaped: six bytes each at worst, then the
    // quotes, a "..." and the NUL.
    const int kValueRenderMax = kValueMax * 6 + 8;

    // Render argument `slot` of a call, given the captured registers.
    // `out` is always null-terminated; empty means "could not be established".
    void DescribeArg(const Slot& s, const Regs& r, char* out, int outSize);

    // `k` is the plan's returnKind -- NOT an assumption that everything
    // returns an LPXLOPER12, which is the defect this signature prevents.
    void DescribeReturn(Kind k, const Regs& r, char* out, int outSize);

    // The XLOPER decoder and its guarded read are internal to xlldecode.cpp:
    // DescribeArg and DescribeReturn are the whole API.
}
