#pragma once
#include "xlltypeplan.h"
#include "xllregs.h"
#include "core/valuewriter.h"

// Describes a captured argument or return value to a core::ValueWriter. Every pointer came from
// somebody else's add-in and is read under SEH; a value that cannot be read confidently writes
// nothing, never a guess.

namespace xll
{
    // The most characters of one string we read; a longer string ends "...".
    const int kValueMax = 128;

    // Argument slot `s` of a call, given the captured registers.
    void DescribeArg(const Slot& s, const Regs& r, core::ValueWriter& w);

    // `k` is the plan's returnKind: not every function returns an LPXLOPER12.
    void DescribeReturn(Kind k, const Regs& r, core::ValueWriter& w);

    // The XLOPER decoder and its guarded read are internal to xlldecode.cpp:
    // DescribeArg and DescribeReturn are the whole API.
}
