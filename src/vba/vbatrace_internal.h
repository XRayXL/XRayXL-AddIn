#pragma once
#include <windows.h>
#include <cstdint>
#include <ostream>
#include "vbaproctable.h"

// The tracer's counters as the report sees them, formatted at disarm after the hooks drain.
namespace vba
{
    constexpr int kMaxDepth = 256;      // deeper than any sane VBA stack

    // Lock-free (key, subkey) -> count. Full is silent: a diagnostic that could stall a VBA
    // statement is worse than one that stops counting.
    struct SeenTable
    {
        static constexpr int kMax = 16;
        // Stored as opcode + 1, so opcode 0 is distinct from an empty slot.
        volatile LONG key[kMax]   = {};
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

    // Exit opcodes with no return-kind mapping, by value so the gap can be named.
    SeenTable& UnmappedExitOps();
    // (store opcode, operand) on untyped returns, so the mapping grows from evidence.
    SeenTable& RetStores();
    // Every exit opcode, so an unknown mid-procedure exit can be named (vbaboundary.h).
    SeenTable& ExitOpsSeen();
    // Read-only by convention: entries are only added and counters only rise.

    ProcTable& Procs();
}
