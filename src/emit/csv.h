#pragma once
#include <windows.h>
#include <string>
#include <cstddef>
#include "rowcsv.h"     // csv::Row and the CSV formatter (Fragment, kHeader, kFragMax)

namespace emit
{

// The trace file, in one of two modes chosen by the buffer size.
//
// SIZE 0 -- SYNCHRONOUS. Formatted and written inside the hook, under a lock.
// Kept only for crash-hunting: every row is on the OS's side of the fence
// before the call returns, so even an uncatchable __fastfail leaves the trace
// complete up to the fault. The cost is that one critical section serialises
// every multithreaded-calc worker for the duration of a syscall -- a profiler
// measuring through its own lock.
//
// SIZE N -- A RING, the production path. The hot path reserves bytes in a
// lock-free MPSC byte ring (no lock, no syscall) and a drain thread stamps the
// seq, batches and writes. Rows are packed variable-length, so the whole budget
// is usable. A process crash loses the in-flight ring; that is the trade size 0
// exists to avoid.
//
// WHEN THE RING FILLS (pauseOnFull): DROP loses the row rather than ever
// stalling the traced thread; PAUSE waits, so a burst the drain cannot match
// throttles the calc to the drain's rate. Either way the count is reported OUT
// OF BAND -- disarm's return value, the status summary, the disarm log line --
// and never in the file, so a lossy CSV reads as complete on its own.
//
// DISARM IS A FENCE: Close() drains and flushes before returning, so a read
// after disarm is complete at any size.

namespace csv
{
    // Opens the trace and writes the header. csv owns the file, so it owns the
    // NAME -- no caller passes a path.
    //
    // Counted: each armed source opens and closes; only the last close tears down.
    //
    // `bufferBytes` 0 is synchronous; N rounds DOWN to a power-of-two byte
    // capacity, reported by RingCapacityBytes(). False only if a NEW file could
    // not be created. The settings of the FIRST opener win: the second is
    // joining a session in progress, not configuring one.
    bool Open(std::size_t bufferBytes, bool pauseOnFull);
    void Close();
    bool IsOpen();      // armed, not "the file exists" -- creation is lazy

    // What was written. Empty until the first record creates the file;
    // survives Close so the disarm report can still name it. Returns a COPY --
    // the drain may set it as it creates the file.
    std::wstring Path();

    long long RowsWritten();

    // Capacity 0 means synchronous. Drops and pauses are cumulative for the
    // session; a record too big for the ring is always a drop, even under PAUSE.
    std::size_t RingCapacityBytes();
    long long   RingDrops();
    long long   RingPauses();

    // One span sequence for both sources: increasing, unique for the life of the process.
    unsigned long long NextSpan();

    // csv::Row and the CSV format live in rowcsv.h; re-exported here so
    // callers still reach `csv::Row` through csv.h.
    //
    // `seq` and `input` are NOT Row fields -- they are prefixes stamped at
    // write: `seq` by the writer (dense, file order), `input` by the
    // producer at emit, so its holes are where rows were dropped.
    // WriteRow materialises the row before returning; nothing in Row outlives
    // the call.
    void WriteRow(const Row& row);

    // Frees this thread's row scratch. Called as a thread ends, never at process exit.
    void ReleaseThreadScratch();

    // After a contained fault: leaves the file's lock if this thread held it.
    void ReleaseHeldByThisThread();

    // XRAYXL_DIAG instrument: faults while holding the file's lock (XRayXL_FaultProbe "CSVLOCK").
    void FaultWhileLockedForProbe();
}
}   // namespace emit
