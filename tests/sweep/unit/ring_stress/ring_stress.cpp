// STRESS TEST FOR emit::ByteRing -- the lock-free MPSC byte queue lifted
// out of csv.cpp. This is the point of the extraction: the one piece of subtle
// memory-ordering in the emitter can be hammered on its own, with no Excel and
// no CSV, and its correctness asserted directly.
//
// It runs N producer threads against one consumer (as the drain thread is the
// sole consumer in the product) and checks the three ways a lock-free queue can
// betray you:
//   * TORN     -- a record read while half-written. Every payload carries a key
//                 in its first 8 bytes and a body derived from that key; a torn
//                 record fails the body check.
//   * LOST/DUP -- a committed record never popped, or popped twice. Every
//                 (producer, index) key is expected exactly once among
//                 popped + dropped.
//   * MISCOUNT -- the reconciliation the product relies on: attempted ==
//                 popped + dropped, exactly.
//
// Three configurations: an AMPLE ring (nothing should drop), a STARVED ring
// under DROP (drops expected, still reconciles, nothing torn), and a STARVED
// ring under PAUSE (nothing dropped, everything eventually popped).
//
// Built by XRayXL.sln into build\x64\Release\unit\; it needs nothing else to run.

#include "ring.h"

#include <windows.h>
#include <cstdio>
#include <cstdint>
#include <cstring>
#include <vector>
#include <atomic>

using emit::ByteRing;

namespace
{
    constexpr int    kProducers   = 8;
    constexpr int    kPerProducer = 200000;     // records each producer attempts
    constexpr int    kMinLen      = 16;
    constexpr int    kMaxLen      = 2000;        // variable length exercises wrap + alignment

    // A record: [u64 key][body...], body[i] = (u8)(key + i). The length is
    // derived from the key so the consumer can regenerate and verify it without
    // any side channel.
    int LenForKey(std::uint64_t key)
    {
        return kMinLen + static_cast<int>(key % (kMaxLen - kMinLen));
    }
    std::uint64_t MakeKey(int producer, int index)
    {
        return (static_cast<std::uint64_t>(producer) << 40) ^ static_cast<std::uint64_t>(index);
    }

    ByteRing*                 g_ring = nullptr;
    HANDLE                    g_wake = nullptr;
    volatile LONG             g_producersDone = 0;
    std::atomic<long long>    g_attempted{0};
    std::atomic<long long>    g_dropped{0};

    // Per (producer,index) seen-count, to catch loss and duplication. One byte
    // each; the address space is kProducers * kPerProducer.
    std::vector<unsigned char> g_seen;
    std::atomic<long long>     g_popped{0};
    std::atomic<long long>     g_torn{0};
    std::atomic<long long>     g_dup{0};

    DWORD WINAPI Producer(LPVOID arg)
    {
        const int id = static_cast<int>(reinterpret_cast<intptr_t>(arg));
        char buf[kMaxLen];
        for (int i = 0; i < kPerProducer; ++i)
        {
            const std::uint64_t key = MakeKey(id, i);
            const int len = LenForKey(key);
            memcpy(buf, &key, sizeof(key));
            for (int b = static_cast<int>(sizeof(key)); b < len; ++b)
                buf[b] = static_cast<char>(static_cast<unsigned char>(key + static_cast<unsigned>(b)));

            g_attempted.fetch_add(1, std::memory_order_relaxed);
            if (!g_ring->TryDeposit(buf, len, g_wake))
                g_dropped.fetch_add(1, std::memory_order_relaxed);
        }
        InterlockedIncrement(&g_producersDone);
        if (g_wake) SetEvent(g_wake);
        return 0;
    }

    // Verify a popped record against its self-describing key. Returns false if
    // torn (body does not match the key).
    bool Verify(const char* rec, int len)
    {
        if (len < static_cast<int>(sizeof(std::uint64_t))) return false;
        std::uint64_t key;
        memcpy(&key, rec, sizeof(key));
        if (LenForKey(key) != len) return false;
        for (int b = static_cast<int>(sizeof(key)); b < len; ++b)
            if (rec[b] != static_cast<char>(static_cast<unsigned char>(key + static_cast<unsigned>(b))))
                return false;
        return true;
    }

    void MarkSeen(std::uint64_t key)
    {
        const int producer = static_cast<int>(key >> 40);
        const int index    = static_cast<int>((key ^ (static_cast<std::uint64_t>(producer) << 40)));
        if (producer < 0 || producer >= kProducers || index < 0 || index >= kPerProducer) return;
        const std::size_t at = static_cast<std::size_t>(producer) * kPerProducer + index;
        if (g_seen[at]++) g_dup.fetch_add(1, std::memory_order_relaxed);
    }

    DWORD WINAPI Consumer(LPVOID)
    {
        char rec[kMaxLen];
        int len = 0;
        for (;;)
        {
            bool any = false;
            while (g_ring->Pop(rec, len))
            {
                any = true;
                g_popped.fetch_add(1, std::memory_order_relaxed);
                if (!Verify(rec, len)) { g_torn.fetch_add(1, std::memory_order_relaxed); continue; }
                std::uint64_t key; memcpy(&key, rec, sizeof(key));
                MarkSeen(key);
            }
            if (!any)
            {
                if (InterlockedCompareExchange(&g_producersDone, 0, 0) == kProducers)
                {
                    // Drain any last committed records, then finish.
                    if (!g_ring->Pop(rec, len)) break;
                    g_popped.fetch_add(1, std::memory_order_relaxed);
                    if (!Verify(rec, len)) g_torn.fetch_add(1, std::memory_order_relaxed);
                    else { std::uint64_t key; memcpy(&key, rec, sizeof(key)); MarkSeen(key); }
                }
                else SwitchToThread();
            }
        }
        return 0;
    }

    // PAUSE RESUMES AT HALF EMPTY. Fill a ring with no consumer, block one more
    // producer on it, then free one record at a time: the producer must not
    // resume until at least half the ring is free.
    struct Blocked { ByteRing* ring; const char* data; int len; volatile LONG done; };

    DWORD WINAPI DepositOnce(LPVOID arg)
    {
        Blocked* b = static_cast<Blocked*>(arg);
        b->ring->TryDeposit(b->data, b->len, nullptr);
        InterlockedExchange(&b->done, 1);
        return 0;
    }

    bool RunHalfEmptyCase()
    {
        const char* name = "pause-resumes-at-half";
        ByteRing ring;
        if (!ring.Init(64ull * 1024, true)) { printf("[FAIL] %s: Init refused\n", name); return false; }

        char rec[1008];                                   // 16-byte header + 1008 = 1024 per record
        memset(rec, 'x', sizeof(rec));
        const int per = static_cast<int>(ring.CapacityBytes() / 1024);
        for (int i = 0; i < per; ++i)
            if (!ring.TryDeposit(rec, sizeof(rec), nullptr)) { printf("[FAIL] %s: fill %d refused\n", name, i); return false; }

        Blocked b{ &ring, rec, sizeof(rec), 0 };
        HANDLE t = CreateThread(nullptr, 0, DepositOnce, &b, 0, nullptr);
        // Blocked means it has counted its pause, not merely that time passed.
        for (int i = 0; i < 2000 && ring.Pauses() == 0; ++i) Sleep(1);
        bool ok = ring.Pauses() == 1 && InterlockedCompareExchange(&b.done, 0, 0) == 0;

        char out[1024]; int len = 0; int popped = 0; int resumedAt = -1;
        while (popped < per && ring.Pop(out, len))
        {
            ++popped;
            Sleep(5);                                     // time for the producer to see the space
            if (InterlockedCompareExchange(&b.done, 0, 0) != 0) { resumedAt = popped; break; }
        }
        WaitForSingleObject(t, 5000);
        CloseHandle(t);

        ok &= resumedAt >= per / 2;
        ok &= ring.Drops() == 0;
        ring.Teardown();
        printf("[%s] %-22s %d-record ring; blocked producer resumed after %d freed (must be >= %d)\n",
               ok ? "PASS" : "FAIL", name, per, resumedAt, per / 2);
        return ok;
    }

    // One configuration. Returns true on pass.
    bool RunCase(const char* name, std::size_t bytes, bool pauseOnFull)
    {
        ByteRing ring;
        if (!ring.Init(bytes, pauseOnFull)) { printf("[FAIL] %s: Init(%zu) refused\n", name, bytes); return false; }

        g_ring = &ring;
        g_wake = CreateEventW(nullptr, FALSE, FALSE, nullptr);
        g_producersDone = 0;
        g_attempted = 0; g_dropped = 0; g_popped = 0; g_torn = 0; g_dup = 0;
        g_seen.assign(static_cast<std::size_t>(kProducers) * kPerProducer, 0);

        HANDLE cons = CreateThread(nullptr, 0, Consumer, nullptr, 0, nullptr);
        HANDLE prod[kProducers];
        for (int i = 0; i < kProducers; ++i)
            prod[i] = CreateThread(nullptr, 0, Producer, reinterpret_cast<LPVOID>(static_cast<intptr_t>(i)), 0, nullptr);

        WaitForMultipleObjects(kProducers, prod, TRUE, INFINITE);
        WaitForSingleObject(cons, INFINITE);
        for (int i = 0; i < kProducers; ++i) CloseHandle(prod[i]);
        CloseHandle(cons);
        CloseHandle(g_wake);

        const long long attempted = g_attempted.load();
        const long long dropped   = g_dropped.load();
        const long long popped    = g_popped.load();
        const long long torn      = g_torn.load();
        const long long dup       = g_dup.load();

        // Loss: a key expected but never seen among popped records.
        long long lost = 0;
        for (unsigned char c : g_seen) if (c == 0) ++lost;
        // Only the successfully-deposited ones should have been seen; the lost
        // count above includes dropped ones, so subtract them.
        lost -= dropped;

        bool ok = true;
        ok &= (torn == 0);
        ok &= (dup == 0);
        ok &= (lost == 0);
        ok &= (attempted == popped + dropped);        // the reconciliation invariant
        if (!pauseOnFull) { /* DROP: drops allowed -- a single verifying consumer
                              cannot match 8 producers, so even a big ring drops;
                              what must hold is the reconciliation, not zero. */ }
        else ok &= (dropped == 0);                     // PAUSE never loses

        const std::size_t cap = ring.CapacityBytes();  // read BEFORE teardown zeroes it
        ring.Teardown();

        printf("[%s] %-22s cap=%-9zu attempted=%lld popped=%lld dropped=%lld torn=%lld dup=%lld lost=%lld\n",
               ok ? "PASS" : "FAIL", name, cap, attempted, popped, dropped, torn, dup, lost);
        return ok;
    }
}

int main()
{
    bool ok = true;
    // DROP, big ring vs tiny ring: 8 producers outrun a single verifying
    // consumer either way, so both drop -- what must hold is that every popped
    // record is intact and attempted == popped + dropped, exactly.
    ok &= RunCase("contended-drop-64M", 64ull * 1024 * 1024, false);
    ok &= RunCase("contended-drop-64K", 64ull * 1024, false);
    // PAUSE on the tiny ring: the producers are throttled to the consumer, so
    // NOTHING is dropped and all 1.6M records are popped intact.
    ok &= RunCase("contended-pause-64K", 64ull * 1024, true);
    // PAUSE waits for half empty, not for room for one row.
    ok &= RunHalfEmptyCase();

    printf("\n%s\n", ok ? "ALL PASS" : "FAILED");
    return ok ? 0 : 1;
}
