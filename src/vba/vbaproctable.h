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

    // THE PROCEDURE TABLE -- a fixed, open-addressed map from p-code
    // trailer to Proc, kept apart from the hot-path hooks so hash/probe/insert
    // can be unit-tested for dedup, collision and full. The caller increments
    // the per-Proc counters through the returned Proc*; name RESOLUTION and the
    // end-of-session REPORT stay in vbatrace.
    //
    // Committed up front and paged in as procedures are seen. Lock-free: a losing
    // racer on an empty-slot claim just probes on.
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
        // 16384, NOT 2048. A real VBA project outgrows 2,048 distinct
        // procedures in one arming session, and the failure was silent in the
        // way that matters: rows kept coming, with no module and no function
        // name, while `named`/`unnamed` -- which count TABLE SLOTS -- both
        // looked healthy. Measured on a 6,723-procedure project built from
        // real-world signatures, where one arm reported `tableFull=14208` and
        // 4,725 procedures that no row ever named.
        //
        // SIZED FOR THE LOAD FACTOR, NOT THE COUNT. This is open addressing
        // with a 64-probe ceiling, so it starts refusing inserts well before it
        // is full: at 8192 the same 6,723 procedures still lost 54 frame
        // pushes, an 82% load. 16384 puts that project at 41%, and the run
        // measures 0. A table sized to just fit is a table that drops.
        //
        // Pages come in as procedures are recorded, and Reset hands them back
        // rather than writing zeros over all of them, so a session tracing 500
        // procedures pays for 500, not for the headroom.
        static constexpr int kSize = 16384;   // power of two, open addressed
        static std::uint64_t Hash(std::uint64_t k);

        Proc*         m_table = nullptr;
        volatile LONG m_count = 0;
    };
}
