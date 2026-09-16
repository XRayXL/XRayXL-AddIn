#include "ring.h"
#include <cstdlib>
#include <cstring>

namespace emit
{
    namespace
    {
        constexpr std::size_t kHdr = 16;        // commitTag + len + pad
        // A floor of last resort, so a degenerate budget falls back to
        // synchronous rather than becoming a one-record ring. The command
        // surface refuses anything below 16 KB well before this.
        constexpr std::size_t kMinBytes = 4096;

        std::size_t Align16(std::size_t n) { return (n + 15u) & ~std::size_t(15u); }

        std::size_t FloorPow2(std::size_t n)
        {
            if (n == 0) return 0;
            std::size_t p = 1;
            while ((p << 1) != 0 && (p << 1) <= n) p <<= 1;
            return p;
        }

        LONG64 LoadAcq(volatile LONG64* p) { return InterlockedCompareExchange64(p, 0, 0); }
    }

    bool ByteRing::Init(std::size_t bytes, bool pauseOnFull)
    {
        Teardown();
        const std::size_t cap = FloorPow2(bytes);
        if (cap < kMinBytes) return false;

        m_buf = static_cast<char*>(malloc(cap));
        if (!m_buf) return false;

        // 0xFF so no untouched location's commitTag equals a real (>= 0) position
        // before the first producer writes it.
        memset(m_buf, 0xFF, cap);
        m_cap = cap;
        m_mask = cap - 1;
        m_tail = m_headPub = 0;
        m_head = 0;
        m_drops = m_pauses = 0;
        m_refilling = 0;
        m_pauseOnFull = pauseOnFull;
        return true;
    }

    void ByteRing::Teardown()
    {
        if (m_buf) { free(m_buf); m_buf = nullptr; }
        m_cap = m_mask = 0;
        m_tail = m_headPub = m_head = 0;
        // m_drops / m_pauses are NOT reset here: the summary reads them after
        // disarm, once the ring is gone. Init resets them for the next session.
    }

    void ByteRing::PutWrapped(LONG64 at, const char* src, int n)
    {
        const std::size_t start = static_cast<std::size_t>(at) & m_mask;
        const std::size_t first = m_cap - start;
        if (static_cast<std::size_t>(n) <= first) { memcpy(m_buf + start, src, static_cast<std::size_t>(n)); return; }
        memcpy(m_buf + start, src, first);
        memcpy(m_buf, src + first, static_cast<std::size_t>(n) - first);
    }

    void ByteRing::GetWrapped(LONG64 at, char* dst, int n)
    {
        const std::size_t start = static_cast<std::size_t>(at) & m_mask;
        const std::size_t first = m_cap - start;
        if (static_cast<std::size_t>(n) <= first) { memcpy(dst, m_buf + start, static_cast<std::size_t>(n)); return; }
        memcpy(dst, m_buf + start, first);
        memcpy(dst + first, m_buf, static_cast<std::size_t>(n) - first);
    }

    bool ByteRing::TryDeposit(const char* data, int n, HANDLE wake)
    {
        if (m_cap == 0 || n < 0) return false;
        const std::size_t need = kHdr + Align16(static_cast<std::size_t>(n));
        // Can never fit: a drop, even under PAUSE.
        if (need > m_cap) { InterlockedIncrement64(&m_drops); return false; }

        LONG64 pos;
        bool waited = false;
        int  spins  = 0;
        for (;;)
        {
            // Close can zero the ring while we wait; drop rather than spin forever.
            if (m_cap == 0) { InterlockedIncrement64(&m_drops); return false; }

            pos = LoadAcq(&m_tail);
            const std::size_t used = static_cast<std::size_t>(pos - LoadAcq(&m_headPub));
            const bool fits = need <= m_cap - used;
            // Full. DROP: lose the record, count it, return.
            if (!fits && !m_pauseOnFull) { InterlockedIncrement64(&m_drops); return false; }
            // PAUSE: once the ring fills, every producer waits until the drain has emptied it to
            // half, so the calculation resumes with room for a burst rather than refilling at once.
            if (m_pauseOnFull && (!fits || InterlockedCompareExchange(&m_refilling, 0, 0) != 0))
            {
                if (fits && used <= m_cap / 2)
                    InterlockedExchange(&m_refilling, 0);
                else
                {
                    InterlockedExchange(&m_refilling, 1);
                    if (!waited) { InterlockedIncrement64(&m_pauses); if (wake) SetEvent(wake); waited = true; }
                    // Yield, then sleep, so a drain stalled in a write does not cost a core per waiter.
                    if (++spins < 64) SwitchToThread(); else Sleep(1);
                    continue;
                }
            }
            if (InterlockedCompareExchange64(&m_tail, pos + static_cast<LONG64>(need), pos) == pos)
                break;                                  // reserved [pos, pos+need)
            // lost the race with another producer; retry
        }

        PutWrapped(pos + static_cast<LONG64>(kHdr), data, n);
        char* h = m_buf + (static_cast<std::size_t>(pos) & m_mask);
        *reinterpret_cast<volatile LONG*>(h + 8) = static_cast<LONG>(n);    // len, before the release
        InterlockedExchange64(reinterpret_cast<volatile LONG64*>(h), pos);  // commitTag = pos (release)
        return true;
    }

    bool ByteRing::Pop(char* out, int& len)
    {
        if (m_cap == 0) return false;
        const LONG64 pos = m_head;
        if (pos >= LoadAcq(&m_tail)) return false;                               // nothing reserved past head
        char* h = m_buf + (static_cast<std::size_t>(pos) & m_mask);
        if (LoadAcq(reinterpret_cast<volatile LONG64*>(h)) != pos) return false; // reserved, not yet committed
        const int n = *reinterpret_cast<volatile LONG*>(h + 8);
        GetWrapped(pos + static_cast<LONG64>(kHdr), out, n);
        len = n;
        m_head = pos + static_cast<LONG64>(kHdr + Align16(static_cast<std::size_t>(n)));
        InterlockedExchange64(&m_headPub, m_head);                               // free the space (release)
        return true;
    }

    long long ByteRing::Drops()  const { return LoadAcq(const_cast<volatile LONG64*>(&m_drops)); }
    long long ByteRing::Pauses() const { return LoadAcq(const_cast<volatile LONG64*>(&m_pauses)); }
    void      ByteRing::ResetCounts() { InterlockedExchange64(&m_drops, 0); InterlockedExchange64(&m_pauses, 0); }
}
