#include "vbaproctable.h"
#include <cstring>

namespace vba
{
    std::uint64_t ProcTable::Hash(std::uint64_t k)
    {
        k ^= k >> 33; k *= 0xff51afd7ed558ccdull;
        k ^= k >> 33; k *= 0xc4ceb9fe1a85ec53ull;
        k ^= k >> 33;
        return k;
    }

    ProcTable::ProcTable()
    {
        m_table = static_cast<Proc*>(VirtualAlloc(nullptr, sizeof(Proc) * kSize,
                                                  MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE));
    }

    Proc* ProcTable::FindOrInsert(std::uint64_t trailer)
    {
        if (trailer == 0 || !m_table) return nullptr;
        std::uint64_t i = Hash(trailer) & (kSize - 1);
        for (int probe = 0; probe < 64; ++probe)
        {
            Proc& p = m_table[i];
            const std::uint64_t seen = p.trailer;
            if (seen == trailer) return &p;
            if (seen == 0)
            {
                // Claim it. A losing racer just probes on.
                if (InterlockedCompareExchange64(
                        reinterpret_cast<volatile LONG64*>(&p.trailer),
                        static_cast<LONG64>(trailer), 0) == 0)
                {
                    InterlockedIncrement(&m_count);
                    return &p;
                }
                if (p.trailer == trailer) return &p;
            }
            i = (i + 1) & (kSize - 1);
        }
        return nullptr;   // full -- the caller counts it in the totals
    }

    void ProcTable::Reset()
    {
        // Pages never touched are already zero; decommitting the rest zeroes them
        // without paging the whole table in.
        if (m_table && m_count != 0)
        {
            const SIZE_T bytes = sizeof(Proc) * kSize;
            VirtualFree(m_table, bytes, MEM_DECOMMIT);
            if (!VirtualAlloc(m_table, bytes, MEM_COMMIT, PAGE_READWRITE)) m_table = nullptr;
        }
        m_count = 0;
    }
}
