#pragma once
#include <windows.h>
#include <cstddef>

// A bounded lock-free MPSC byte queue, kept apart from the CSV formatting and file I/O so it
// can be stress-tested on its own.
//
// Many producers (the calc-thread hooks) deposit variable-length records; one
// consumer (the drain thread) pops them in order. A record is a 16-byte header
// followed by exactly its bytes, wrapping the buffer end:
//
//   +00  commitTag (LONG64)  == the record's own start cursor once COMMITTED;
//                              any other value means "reserved but not written"
//                              (or a stale earlier record left here by a wrap)
//   +08  payloadLen (LONG)
//   +0C  pad -- 16 bytes keeps the header 8-aligned and, since the buffer is a
//              power of two and records are 16-aligned, never straddling the
//              wrap. Only the payload wraps.
//   +10  payload (payloadLen bytes; MAY wrap the end of the buffer)
//
// A producer reserves a range by CAS on the tail, writes the payload and length, then commits
// by storing its own start position as the tag. The consumer reads a header only when the
// position is both reserved (pos < tail) and committed (commitTag == pos), so it never sees a
// half-written record or a stale one from an earlier wrap. Init fills with 0xFF, which no
// position can equal.
//
// TryDeposit signals a caller-supplied wake event when it blocks, so a full ring need not wait
// for the drain's next poll.

namespace emit
{
    class ByteRing
    {
    public:
        ByteRing() = default;
        ByteRing(const ByteRing&) = delete;
        ByteRing& operator=(const ByteRing&) = delete;

        // ~`bytes`, rounded DOWN to a power of two. False (and inactive) when
        // the budget is too small or the allocation fails, so the caller falls
        // back to the synchronous path. `pauseOnFull`: false = DROP, true =
        // PAUSE.
        bool  Init(std::size_t bytes, bool pauseOnFull);

        // Free the buffer and reset to inactive. Idempotent. The caller must
        // have already stopped the drain thread -- Teardown does not.
        void  Teardown();

        bool        Active()        const { return m_cap != 0; }
        std::size_t CapacityBytes() const { return m_cap; }

        // Producer, any thread. Copies `n` bytes as one record. On a full ring DROP counts a
        // drop and returns false; PAUSE signals `wake` and waits until the drain has emptied
        // the ring to half. Returns false without blocking when `n` can never fit or the ring
        // is inactive.
        bool  TryDeposit(const char* data, int n, HANDLE wake);

        // CONSUMER, single thread only. The length of the next committed record, or -1 when
        // none is ready, so the caller can size the buffer it pops into.
        int   PendingLen() const;

        // CONSUMER, single thread only. Pops the next committed record into
        // `out` -- which must be large enough for the biggest record ever
        // deposited -- sets `len`, and returns true. False when nothing is ready
        // (empty, or the next record's producer has not committed yet).
        bool  Pop(char* out, int& len);

        // Cumulative since the last ResetCounts or Init. A record too big for the ring is a drop even under PAUSE.
        long long Drops()  const;
        long long Pauses() const;
        // Survives Teardown, so a synchronous session must call this or it reports the last ring's loss.
        void      ResetCounts();

    private:
        void PutWrapped(LONG64 at, const char* src, int n);
        void GetWrapped(LONG64 at, char* dst, int n);

        char*            m_buf      = nullptr;   // the byte buffer
        std::size_t      m_cap      = 0;         // power of two, or 0 = inactive
        std::size_t      m_mask     = 0;
        volatile LONG64  m_tail     = 0;         // producer reservation cursor
        volatile LONG64  m_headPub  = 0;         // bytes the consumer has freed
        LONG64           m_head     = 0;         // consumer cursor (drain only)
        volatile LONG64  m_drops    = 0;         // DROP policy: rows lost
        volatile LONG64  m_pauses   = 0;         // PAUSE policy: rows that waited
        volatile LONG    m_refilling = 0;        // PAUSE policy: set when full, cleared at half empty
        bool             m_pauseOnFull = false;
    };
}
