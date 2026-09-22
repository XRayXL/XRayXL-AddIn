#pragma once
#include <windows.h>
#include <string>
#include <cstddef>
#include "rowcsv.h"     // csv::Row and the CSV formatter (Fragment, kHeader)
#include "core/tracemodes.h"

namespace emit
{

// The trace file: CSV or JSON Lines, chosen at Open, and written in one of two modes chosen by
// the buffer size.
//
// Size 0 is synchronous: written inside the hook under a lock, so even a __fastfail leaves the
// trace complete up to the fault. It serialises every calc worker, so it is for crash-hunting
// only.
//
// Size N is a ring: the hot path reserves bytes in a lock-free MPSC ring and a drain thread
// stamps the seq, batches and writes. A process crash loses the in-flight ring.
//
// When the ring fills, DROP loses the row and PAUSE waits for the drain. The count is reported
// out of band, never in the file. Close() drains and flushes before returning.

namespace csv
{
    // Opens the trace and writes the header. Counted: each armed source opens and closes, and
    // only the last close tears down. The first opener's settings win.
    //
    // `bufferBytes` 0 is synchronous; N rounds down to a power of two, reported by
    // RingCapacityBytes(). `format` is latched for the file, values included, and so is `breaks`,
    // the optional column. False only if a new file could not be created.
    bool Open(std::size_t bufferBytes, bool pauseOnFull, core::modes::Format format, bool breaks);
    void Close();
    bool IsOpen();      // armed, not "the file exists" -- creation is lazy
    bool HasBreaksColumn();     // the optional `breaks` column, latched at Open

    // The trace's name, decided at Open although the file is created lazily. A copy: the drain may set it.
    std::wstring Path();

    // Whether the file has been created yet; Path() names it either way.
    bool FileExists();

    long long RowsWritten();

    // Capacity 0 means synchronous. Drops and pauses are cumulative for the
    // session; a record too big for the ring is always a drop, even under PAUSE.
    std::size_t RingCapacityBytes();
    long long   RingDrops();
    long long   RingPauses();

    // One span sequence for both sources: increasing, unique for the life of the process.
    unsigned long long NextSpan();

    // `seq` and `input` are not Row fields: they are prefixes stamped at write, `seq` by the
    // writer and `input` by the producer, so its holes are where rows were dropped. WriteRow
    // materialises the row before returning.
    void WriteRow(const Row& row);

    // Frees this thread's row scratch. Called as a thread ends, never at process exit.
    void ReleaseThreadScratch();

    // After a contained fault: leaves the file's lock if this thread held it.
    void ReleaseHeldByThisThread();

    // XRAYXL_DIAG instrument: faults while holding the file's lock (XRayXL_FaultProbe "CSVLOCK").
    void FaultWhileLockedForProbe();
}
}   // namespace emit
