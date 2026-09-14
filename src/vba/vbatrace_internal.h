#pragma once
#include <windows.h>
#include <cstdint>
#include <ostream>
#include "vbaproctable.h"

// THE TRACER'S COUNTERS AS THE REPORT SEES THEM. Internal to vba/: vbatrace.cpp
// owns the state, vbareport.cpp formats it at disarm after the hooks have
// drained.
namespace vba
{
    constexpr int kMaxDepth = 256;      // deeper than any sane VBA stack

    // A SMALL LOCK-FREE (key, subkey) -> COUNT TABLE for the three "which
    // opcode, how often" diagnostics below. Full is SILENT: a diagnostic
    // that could stall a VBA statement would be worse than one that stops
    // counting.
    struct SeenTable
    {
        static constexpr int kMax = 16;
        // Keys are stored as opcode + 1, so opcode 0 is distinct from an empty slot.
        volatile LONG key[kMax]   = {};      // 0 empty, otherwise opcode + 1
        volatile LONG sub[kMax]   = {};
        volatile LONG count[kMax] = {};

        void Note(std::uint16_t k, std::int32_t sk = 0)
        {
            const LONG want = static_cast<LONG>(k) + 1;
            for (int i = 0; i < kMax; ++i)
            {
                const LONG o = InterlockedCompareExchange(&key[i], 0, 0);
                if (o == want && InterlockedCompareExchange(&sub[i], 0, 0) == sk)
                { InterlockedIncrement(&count[i]); return; }
                if (o == 0 && InterlockedCompareExchange(&key[i], want, 0) == 0)
                {
                    InterlockedExchange(&sub[i], sk);
                    InterlockedIncrement(&count[i]);
                    return;
                }
            }
        }
        void Reset()
        {
            for (int i = 0; i < kMax; ++i)
            {
                InterlockedExchange(&key[i], 0);
                InterlockedExchange(&sub[i], 0);
                InterlockedExchange(&count[i], 0);
            }
        }
        // " label[key]=count" per entry, or " label[key@sub]=count".
        // Not const: the interlocked reads take a non-const pointer.
        void Print(std::ostream& o, const char* label, bool withSub)
        {
            for (int i = 0; i < kMax; ++i)
            {
                const LONG k = InterlockedCompareExchange(&key[i], 0, 0);
                if (!k) continue;
                o << " " << label << "[" << (k - 1);      // stored biased by one
                if (withSub) o << "@" << InterlockedCompareExchange(&sub[i], 0, 0);
                o << "]=" << InterlockedCompareExchange(&count[i], 0, 0);
            }
        }
    };

    // ---- EXIT OPCODES WE CANNOT NAME A RETURN KIND FOR --------------
    // Recorded by VALUE, because the value is the thing that has to be
    // mapped; a count alone would say a gap exists without saying which.
    SeenTable& UnmappedExitOps();
    // The (store opcode, operand) pairs seen on procedures whose return could
    // not be typed, so the mapping is extended from evidence.
    SeenTable& RetStores();
    // EVERY exit opcode seen, mapped or not: how an unknown mid-procedure
    // exit construct gets named (vbaboundary.h).
    SeenTable& ExitOpsSeen();
    // The procedure table, read-only by convention: entries are only added
    // and counters only rise.
    ProcTable& Procs();
}
