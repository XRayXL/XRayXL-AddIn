#pragma once
#include <windows.h>
#include <cstdint>

namespace vba
{
    struct Proc
    {
        volatile std::uint64_t trailer;   // 0 = empty
        volatile std::uint64_t calls;
        volatile std::uint64_t statements;
        volatile std::uint64_t ticks;
        volatile std::uint64_t maxDepth;

        // Re-resolved on every push, never cached: a recompile can hand a freed trailer to a
        // different procedure.
        char qualModule[384];             // "[Book.xlsm]Project.Module"
        char function[256];               // "MyProcedure"

        // One publisher at a time: two threads write the same name, but can still tear it if an
        // edit changed its length mid-session.
        volatile LONG publishing;
    };

    // Trailer -> Proc, lock-free and open-addressed; kept apart from the hooks so it can be
    // unit-tested.
    class ProcTable
    {
    public:
        ProcTable();

        // nullptr when full (64 probes exhausted) or for trailer 0, which marks an empty slot.
        Proc* FindOrInsert(std::uint64_t trailer);

        // An unused slot has trailer == 0; nullptr only for an out-of-range index.
        Proc* At(int i) { return (m_table && i >= 0 && i < kSize) ? &m_table[i] : nullptr; }
        int   Size() const { return kSize; }
        long  Count() const { return m_count; }   // distinct procedures claimed

        // At re-arm, so counts belong to the arming session rather than the process.
        void  Reset();

    private:
        // Sized for the load factor: a 64-probe ceiling refuses inserts well before full. Pages
        // come in as used and Reset hands them back, so a small session pays little.
        static constexpr int kSize = 16384;   // power of two

        static std::uint64_t Hash(std::uint64_t k);

        Proc*         m_table = nullptr;
        volatile LONG m_count = 0;
    };
}
