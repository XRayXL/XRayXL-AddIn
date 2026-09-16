// UNIT TEST for vba::ProcTable (vba/vbaproctable.cpp) -- the open-addressed
// procedure table lifted out of vbatrace.cpp. Keyed by an opaque uint64
// trailer, so it needs no Excel and no memory: insert trailers, assert dedup,
// distinct slots, survival of collisions, capacity, and Reset.
//
// Built by XRayXL.sln into build\x64\Release\unit\.

#include "vbaproctable.h"

#include <cstdio>
#include <cstdint>
#include <vector>

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }
}

int main()
{
    static vba::ProcTable t;   // ~1.3 MB; keep off the stack

    // ---- a new trailer is inserted; the same trailer returns the same slot ---
    {
        vba::Proc* a = t.FindOrInsert(0x1000);
        vba::Proc* b = t.FindOrInsert(0x1000);
        Check(a != nullptr && a == b, "same trailer -> same Proc");
        Check(t.Count() == 1, "one distinct procedure counted");
        // the per-Proc counters are the caller's to drive, through the pointer
        a->calls = 7;
        Check(t.FindOrInsert(0x1000)->calls == 7, "counters persist on the slot");
    }

    // ---- distinct trailers get distinct slots, count rises -------------------
    {
        vba::Proc* a = t.FindOrInsert(0x2000);
        vba::Proc* b = t.FindOrInsert(0x3000);
        Check(a && b && a != b, "distinct trailers -> distinct Procs");
        Check(t.Count() == 3, "three distinct procedures counted");
    }

    // ---- Reset clears everything --------------------------------------------
    {
        t.Reset();
        Check(t.Count() == 0, "Reset zeroes the count");
        bool anyLive = false;
        for (int i = 0; i < t.Size(); ++i) if (t.At(i)->trailer) anyLive = true;
        Check(!anyLive, "Reset zeroes every slot");
        Check(t.FindOrInsert(0x1000)->calls == 0, "a slot reused after Reset starts clean");
        t.Reset();
    }

    // ---- trailer 0 marks an empty slot, so it can never be a key --------------
    {
        Check(t.FindOrInsert(0) == nullptr, "trailer 0 is refused rather than handed an unclaimed slot");
        Check(t.Count() == 0, "refusing trailer 0 claims nothing");
        t.Reset();
    }

    // ---- fill well past a hash bucket: many distinct trailers, all findable --
    {
        const int N = 1500;   // < Size (16384); all must fit and stay distinct
        std::vector<vba::Proc*> got(N);
        for (int i = 0; i < N; ++i) got[i] = t.FindOrInsert(0x100000ull + i);
        bool allInserted = true, allStable = true;
        for (int i = 0; i < N; ++i)
        {
            if (!got[i]) allInserted = false;
            if (t.FindOrInsert(0x100000ull + i) != got[i]) allStable = false;
        }
        Check(allInserted, "1500 distinct trailers all inserted");
        Check(allStable, "every trailer re-finds its own slot (collisions probed correctly)");
        Check(t.Count() == N, "count equals the number of distinct trailers");
    }

    // ---- out-of-range At is null, not a crash --------------------------------
    Check(t.At(-1) == nullptr && t.At(t.Size()) == nullptr, "At() bounds-checks");

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
