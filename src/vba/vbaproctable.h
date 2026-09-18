#pragma once
#include <windows.h>
#include <cstdint>

namespace vba
{
    // ONE PROCEDURE'S ROW in the name/count table.
    struct Proc
    {
        volatile std::uint64_t trailer;   // 0 = empty
        volatile std::uint64_t calls;
        volatile std::uint64_t statements;
        volatile std::uint64_t ticks;
        volatile std::uint64_t maxDepth;

        // Re-resolved on EVERY frame push, never cached: editing VBA recompiles
        // the module, and a freed trailer can be handed to a DIFFERENT
        // procedure. Sized to VBA's own limits.
        char qualModule[384];             // "[Book.xlsm]Project.Module"
        char function[256];               // "MyProcedure"

        // One publisher at a time: two threads pushing frames for the same
        // procedure write the SAME name, but a concurrent pair can still tear it
        // if an edit changed its length mid-session. The READ side is already
        // safe -- Report()/ProcSnapshot run after the hooks have drained.
        volatile LONG publishing;
    };

    // A fixed, open-addressed map from p-code trailer to Proc, kept apart from the hot-path
    // hooks so it can be unit-tested. Committed up front and paged in as procedures are seen.
    // Lock-free: a losing racer on an empty-slot claim just probes on.
    class ProcTable
    {
    public:
        ProcTable();

        // The slot for `trailer`, claiming an empty one if new. nullptr when
        // the table is full (64 probes exhausted), or for trailer 0, which marks
        // an empty slot. The caller counts that.
        Proc* FindOrInsert(std::uint64_t trailer);

        // Slot `i` in [0, Size()) for iteration and snapshots. An unused slot has
        // trailer == 0. nullptr only for an out-of-range index.
        Proc* At(int i) { return (m_table && i >= 0 && i < kSize) ? &m_table[i] : nullptr; }
        int   Size() const { return kSize; }
        long  Count() const { return m_count; }   // distinct procedures claimed

        // Called at re-arm, so counts belong to the arming session rather than
        // to the process.
        void  Reset();

    private:
        // Sized for the load factor, not the count: open addressing with a 64-probe ceiling
        // refuses inserts well before it is full, and a real project can hold several thousand
        // procedures. Pages come in as procedures are recorded and Reset hands them back, so a
        // small session pays only for what it uses.
        static constexpr int kSize = 16384;   // power of two, open addressed
        static std::uint64_t Hash(std::uint64_t k);

        Proc*         m_table = nullptr;
        volatile LONG m_count = 0;
    };
}
