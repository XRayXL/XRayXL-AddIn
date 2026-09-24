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
        // Guards against a bad trailer read.
        constexpr std::uint32_t kMaxProcSize = kTrl_procSizeMax;

        // IsProcTerminatorSlot in vbaderive.cpp must agree with the deny-list below.

        // Every read is of someone else's memory on a VBA thread, so a fault is an answer.
        using core::RdU64;
        using core::RdU16;
        using core::RdI32;
    }

    Activation ActivationStart(std::uint64_t trailer, std::uint64_t savedRegs)
    {
        std::uint64_t rsi = 0;
        std::uint16_t procSize16 = 0;

        if (!RdU64(savedRegs + kReg_rsi, rsi))            return Activation::Unavailable;
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

        // A deny-list: an allow-list would stop closing frames for a procedure kind nobody tested.

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
