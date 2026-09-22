#include "vbaboundary.h"
#include "vbaregs.h"
#include "vbatrailer.h"
#include "vbaslots.h"

#include <windows.h>
#include "core/safemem.h"

namespace vba
{
    namespace
    {
        // kTrl_procSize (trailer+0x0C) is in vbatrailer.h, the one authority.

        // Guards against a bad trailer read. ProcSize is a WORD, so this is its maximum.
        constexpr std::uint32_t kMaxProcSize = kTrl_procSizeMax;

        // The slots in the deny-list below are in vbaslots.h; IsProcTerminatorSlot
        // in vbaderive.cpp must agree with it.

        // Everything below reads memory that belongs to somebody else, on a VBA
        // thread, so every read is guarded and a fault is an answer rather than
        // a crash. SEH lives in leaf functions with no C++ objects in scope --
        // __try cannot share a frame with anything that needs unwinding.
        using core::RdU64;
        using core::RdU16;
        using core::RdI32;
    }

    Activation ActivationStart(std::uint64_t trailer, std::uint64_t savedRegs)
    {
        std::uint64_t rsi = 0;
        std::uint16_t procSize16 = 0;

        if (!RdU64(savedRegs + kReg_rsi, rsi))            return Activation::Unavailable;
        // A WORD, as vbatrailer.h says and as the p-code walk reads it.
        if (!RdU16(trailer + kTrl_procSize, procSize16))  return Activation::Unavailable;
        const std::uint32_t procSize = procSize16;

        // Bounds before arithmetic: `trailer <= procSize` would make the subtraction below
        // wrap, and a boundary that fires on a wrapped pointer is worse than one that declines.
        if (procSize == 0 || procSize > kMaxProcSize)    return Activation::Unavailable;
        if (rsi < 2 || trailer <= procSize)              return Activation::Unavailable;

        // RSI-2 is the instruction being dispatched; trailer-ProcSize is the
        // procedure's first byte. Equal means this is the prologue.
        return (rsi - 2 == trailer - procSize) ? Activation::Yes : Activation::No;
    }

    Ending ProcedureEnd(std::uint64_t savedRegs, std::uint16_t* opOut)
    {
        std::uint64_t rsi = 0;
        std::uint16_t op  = 0;

        if (!RdU64(savedRegs + kReg_rsi, rsi)) return Ending::Unreadable;
        if (rsi < 2)                            return Ending::Unreadable;
        if (!RdU16(rsi - 2, op))               return Ending::Unreadable;
        if (opOut) *opOut = op;

        // A DENY-LIST, not an allow-list, and vbaboundary.h says why: only 5 of
        // the 25 exit slots have ever been observed firing, so listing the ones
        // that DO end a procedure would silently stop closing frames for any
        // procedure kind nobody tested.
        if (op == kSlot_GoSubReturn)  return Ending::No;
        if (op == kSlot_ZeroRetVal)    return Ending::No;
        if (op == kSlot_ZeroRetValVar) return Ending::No;
        return Ending::Yes;
    }

    bool ExitOperand(std::uint64_t savedRegs, std::int32_t& out)
    {
        std::uint64_t rsi = 0;
        return RdU64(savedRegs + kReg_rsi, rsi) && rsi != 0 && RdI32(rsi, out);
    }
}
