#include "csv.h"
#include "ring.h"
#include "core/log.h"
#include "core/excel_api.h"
#include "core/textbuf.h"
#include "core/valueformat.h"
#include "rowjson.h"
#include <cstdio>
#include <cstdlib>
#include <cstring>

namespace emit
{

namespace csv
{
    namespace
    {
        CRITICAL_SECTION g_cs;           // never deleted: the module is pinned for the process
        bool             g_csReady = false;
        // How deep this thread is in g_cs, so a contained fault can leave it (ReleaseHeldByThisThread).
        __declspec(thread) int t_csDepth = 0;
        void Lock()   { EnterCriticalSection(&g_cs); ++t_csDepth; }
        void Unlock() { --t_csDepth; LeaveCriticalSection(&g_cs); }
        // ARMED, not OPEN: the file is created lazily, on the first record.
        // IsOpen() reports this, because the emitter's callers mean "are
        // we tracing", and EmitRow must build rows before the file exists:
        // building one is what causes it to exist.
        volatile LONG    g_prepared = 0;
        bool Prepared() { return InterlockedCompareExchange(&g_prepared, 0, 0) != 0; }
        // Producers between their armed check and leaving WriteRow. Close waits for
        // zero before freeing the ring, so a record never lands in freed memory.
        volatile LONG    g_producers = 0;
        // A close that could not wait out a producer or join the drain leaves the
        // ring to them for the process; later sessions write synchronously.
        bool             g_ringAbandoned = false;
        // Armed sources still producing; the last Close tears down. Guarded by g_cs.
        int              g_owners = 0;
        HANDLE           g_file = INVALID_HANDLE_VALUE;  // INVALID until the first record
        std::wstring     g_path;                          // set when the file is created; read under g_cs
        // The NAME is chosen at arm so arming can report it; the FILE is still
        // made lazily. The id is a timestamp nobody can predict, so naming at
        // the first record meant a session had to finish before its file could
        // be named.
        std::wstring     g_planned;                       // decided at arm; read under g_cs
        volatile LONG64  g_rows = 0;
        // One counter for every source, so a `seq` value appears once in a file.
        volatile LONG64  g_seq  = 0;   // WRITER-stamped: dense, file order
        volatile LONG64  g_in   = 0;   // PRODUCER-stamped at emit: holes = drops
        // Span ids for both sources. Never reset, so a call running across a re-arm cannot reuse one.
        volatile LONG64  g_span = 0;
        // The file's format, latched at Open for the life of the file.
        core::modes::Format g_format = core::modes::Format::Csv;

        bool Jsonl() { return g_format == core::modes::Format::Jsonl; }

        // What the writer puts before a row: `seq,` for CSV, `{"seq":N,` opening a JSON object.
        int SeqPrefix(long long seq, char* out, int cap)
        {
            return _snprintf_s(out, cap, _TRUNCATE, Jsonl() ? "{\"seq\":%lld," : "%lld,", seq);
        }

        void WriteRaw(const char* text, DWORD len)
        {
            if (g_file == INVALID_HANDLE_VALUE) return;
            DWORD written = 0;
            WriteFile(g_file, text, len, &written, nullptr);
        }

        // The one place the trace file is named. The id is a monotonic OS tick, so it rises
        // across arms, processes and reboots; the pid breaks any tie. The path goes to UTF-8
        // properly, since %TEMP% carries a user name that may not be ASCII.
        std::string Narrow(const std::wstring& w)
        {
            if (w.empty()) return std::string();
            const int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, nullptr, 0, nullptr, nullptr);
            if (n <= 1) return std::string();
            std::string s(static_cast<size_t>(n - 1), 0);
            WideCharToMultiByte(CP_UTF8, 0, w.c_str(), -1, &s[0], n, nullptr, nullptr);
            return s;
        }

        std::wstring BuildTracePath()
        {
            FILETIME ft{};
            GetSystemTimePreciseAsFileTime(&ft);
            ULARGE_INTEGER id; id.LowPart = ft.dwLowDateTime; id.HighPart = ft.dwHighDateTime;
            // The same output root as the log and the corpus (XRAYXL_OUTPUT_DIR,
            // else %TEMP%\XRayXL): a second copy of that rule here once sent the
            // traces to %TEMP% while the logs went to the session directory.
            const std::wstring dir = core::EnsureAppSubdir(L"TraceFiles");
            if (dir.empty()) return L"";
            wchar_t name[96];
            swprintf_s(name, LR"(\XRayXL_Trace_%llu_%lu.%s)",
                       id.QuadPart, GetCurrentProcessId(), Jsonl() ? L"jsonl" : L"csv");
            return dir + name;
        }

        // The header, column order, escaping and Fragment() are rowcsv's;
        // this file is the FILE, the ring and the drain.

        // The largest row: the argument and return columns at their limit, every byte a quote
        // that escaping doubles, and room for the rest.
        constexpr std::size_t kMaxRowBytes = 4 * core::kMaxValueBytes + 64 * 1024;

        // Per-thread scratch for the escaped row, grown to the row and kept for the next.
        // WriteRow runs concurrently on every calc worker in ring mode.
        __declspec(thread) core::TextBuf t_frag;   // freed when its thread ends (ReleaseThreadScratch)
        char* FragBuf(std::size_t need)
        {
            t_frag.limit = kMaxRowBytes;
            t_frag.Clear();
            return t_frag.Reserve(need);   // null -- WriteRow then drops the row rather than fault
        }

        // Creates the file on the first record that reaches a writer, never on the arm path, so
        // a session that traces nothing leaves no empty file. Takes g_cs recursively and
        // publishes g_file last, because readers gate on it.
        bool EnsureFile()
        {
            if (g_file != INVALID_HANDLE_VALUE) return true;
            Lock();
            if (g_file == INVALID_HANDLE_VALUE)          // recheck under the lock
            {
                const std::wstring path = g_planned.empty() ? BuildTracePath() : g_planned;
                HANDLE h = CreateFileW(path.c_str(), GENERIC_WRITE, FILE_SHARE_READ,
                                       nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
                if (h != INVALID_HANDLE_VALUE)
                {
                    g_path = path;
                    g_seq  = 0;
                    g_rows = 0;
                    g_file = h;                          // publish before the header write below
                    // JSON Lines has no header: every line names its own keys.
                    if (!Jsonl()) WriteRaw(kHeader, static_cast<DWORD>(strlen(kHeader)));
                }
            }
            const bool ok = (g_file != INVALID_HANDLE_VALUE);
            Unlock();
            // Creation is deliberately not logged: the arm line names this file
            // in full and the disarm line reports what reached it, so a third
            // line about the same file at a third moment is noise.
            return ok;
        }

        // The ring. The byte queue is ring.h; here are the drain thread, its events, the write
        // batch, seq stamping and the WriteFile.
        emit::ByteRing   g_ring;

        // THE DRAIN as one thing -- thread, stop/wake events, and the
        // coalescing batch -- so the code reads `g_drain.wake` rather than six
        // parallel globals.
        struct Drain
        {
            HANDLE      thread = nullptr;
            HANDLE      stop   = nullptr;   // manual-reset: signals the loop to end
            HANDLE      wake   = nullptr;   // auto-reset: a blocked producer kicks the drain
            char*       batch  = nullptr;   // write-coalescing buffer
            std::size_t cap    = 0;
            std::size_t used   = 0;
        };
        Drain g_drain;

        // Flushes first if `line` would not fit. No lock: in ring mode the
        // drain is the sole file writer.
        void BatchAppend(const char* line, std::size_t len)
        {
            if (g_drain.used + len > g_drain.cap) { WriteRaw(g_drain.batch, static_cast<DWORD>(g_drain.used)); g_drain.used = 0; }
            if (len > g_drain.cap) { WriteRaw(line, static_cast<DWORD>(len)); return; }  // paranoia
            memcpy(g_drain.batch + g_drain.used, line, len);
            g_drain.used += len;
        }

        void DrainAll()
        {
            g_drain.used = 0;

            // Grown to the largest record so far: DrainAll runs only on the one drain thread.
            static core::TextBuf frag;
            frag.limit = static_cast<std::size_t>(-1);
            int flen = 0;
            for (int next; (next = g_ring.PendingLen()) >= 0; )
            {
                frag.Clear();
                char* into = frag.Reserve(static_cast<std::size_t>(next));
                if (!into || !g_ring.Pop(into, flen)) break;
                if (!EnsureFile()) return;   // the drain creates the file on the first record
                char seqb[32];
                const int sl = SeqPrefix(static_cast<long long>(InterlockedIncrement64(&g_seq)),
                                         seqb, sizeof seqb);
                BatchAppend(seqb, static_cast<std::size_t>(sl));
                BatchAppend(into, static_cast<std::size_t>(flen));
                InterlockedIncrement64(&g_rows);
            }

            if (g_drain.used) { WriteRaw(g_drain.batch, static_cast<DWORD>(g_drain.used)); g_drain.used = 0; }
        }

        // A high-resolution waitable timer paces the poll at about 1 ms. kernel32 rather than
        // timeBeginPeriod, to add no link dependency; without it (pre-1803) the wait is correct
        // but coarser.
        #ifndef CREATE_WAITABLE_TIMER_HIGH_RESOLUTION
        #define CREATE_WAITABLE_TIMER_HIGH_RESOLUTION 0x00000002
        #endif
        DWORD WINAPI DrainProc(LPVOID)
        {
            HANDLE timer = CreateWaitableTimerExW(nullptr, nullptr,
                CREATE_WAITABLE_TIMER_HIGH_RESOLUTION, TIMER_ALL_ACCESS);
            LARGE_INTEGER due; due.QuadPart = -10000;   // first fire in 1 ms (100 ns units)
            const bool haveTimer = timer && SetWaitableTimer(timer, &due, 1, nullptr, nullptr, FALSE);
            // The wake is what makes PAUSE responsive: a producer that fills
            // the ring kicks the drain rather than stalling until the next tick.
            HANDLE hs[3] = { g_drain.stop, g_drain.wake, timer };
            const DWORD nh = haveTimer ? 3u : 2u;
            for (;;)
            {
                const DWORD w = WaitForMultipleObjects(nh, hs, FALSE, haveTimer ? INFINITE : 1);
                const bool stopping = (w == WAIT_OBJECT_0);
                // A failed wait returns at once; pace it rather than spin a core.
                if (w == WAIT_FAILED) Sleep(1);
                DrainAll();
                if (stopping) break;
            }
            DrainAll();      // belt and braces after the stop was seen
            if (timer) { CancelWaitableTimer(timer); CloseHandle(timer); }
            return 0;
        }

        // Join the drain before freeing its memory. False if it would not join,
        // in which case nothing is freed.
        bool RingTeardown()
        {
            bool joined = true;
            if (g_drain.thread)
            {
                if (g_drain.stop) SetEvent(g_drain.stop);
                // A drain stuck in WriteFile still reads the ring: leak rather than free under it.
                joined = (WaitForSingleObject(g_drain.thread, 5000) == WAIT_OBJECT_0);
                if (joined)
                {
                    CloseHandle(g_drain.thread);
                    g_drain.thread = nullptr;
                }
                else
                {
                    core::Log::Warning(
                        "trace ring: the drain thread did not finish within 5s -- its buffers "
                        "are left allocated rather than freed under it, and this session's "
                        "trace may be incomplete");
                }
            }
            if (!joined) return false;      // the drain still owns everything below

            if (g_drain.stop) { CloseHandle(g_drain.stop); g_drain.stop = nullptr; }
            if (g_drain.wake) { CloseHandle(g_drain.wake); g_drain.wake = nullptr; }
            g_ring.Teardown();
            if (g_drain.batch) { free(g_drain.batch); g_drain.batch = nullptr; }
            g_drain.used = 0;
            return true;
        }

        // False if a producer is still inside WriteRow when the time is up.
        bool WaitForProducers(DWORD ms)
        {
            const ULONGLONG deadline = GetTickCount64() + ms;
            while (InterlockedCompareExchange(&g_producers, 0, 0) != 0)
            {
                if (GetTickCount64() >= deadline) return false;
                Sleep(1);
            }
            return true;
        }
    }

    bool Open(std::size_t bufferBytes, bool pauseOnFull, core::modes::Format format)
    {
        if (!g_csReady) { InitializeCriticalSection(&g_cs); g_csReady = true; }
        Lock();
        // PREPARE ONLY -- no file is created here. Idempotent: whichever
        // source arms first sets up the ring and drain, and the other's call is
        // a no-op that keeps it rather than tearing it down.
        if (Prepared()) { g_owners++; Unlock(); return true; }
        // A drain that never joined, or a producer never counted out, still owns the ring; write synchronously instead.
        const bool ringFree = !g_ringAbandoned && RingTeardown();
        // Every new session, ring or not: a synchronous one otherwise reports the last ring's drops.
        g_ring.ResetCounts();
        g_rows = 0;
        g_seq  = 0;
        g_in   = 0;             // reset HERE (arm), before any producer stamps it
        g_path.clear();
        g_format = format;
        core::LatchFormat(format);      // every value in this file is spelt one way
        g_planned = BuildTracePath();   // named now, created on the first record

        // Logged HERE rather than at the arm sites, because the name belongs to
        // this module and either arm path can be the only one that runs --
        // a workbook with macros and no add-ins arms VBA alone. Open being
        // idempotent makes that one line per arm, wherever the arm came from.
        {
            const std::string p = Narrow(g_planned);
            core::Log::Info(p.empty() ? std::string("trace file: could not be named")
                                : std::string("trace file (armed): ") + p);
        }
        // A ring that stands up but whose batch or drain thread cannot be
        // created falls back to synchronous -- a performance choice, not a
        // correctness one.
        if (ringFree && g_ring.Init(bufferBytes, pauseOnFull))
        {
            g_drain.batch = static_cast<char*>(malloc(1u << 18));   // 256 KB coalesce
            if (g_drain.batch)
            {
                g_drain.cap = 1u << 18;
                g_drain.stop   = CreateEventW(nullptr, TRUE,  FALSE, nullptr);  // manual: stays set
                g_drain.wake   = CreateEventW(nullptr, FALSE, FALSE, nullptr);  // auto: one kick per set
                // Without both events the drain loop would spin; stay synchronous.
                if (g_drain.stop && g_drain.wake)
                    g_drain.thread = CreateThread(nullptr, 0, DrainProc, nullptr, 0, nullptr);
                if (!g_drain.thread) RingTeardown();   // no thread -> stay sync
            }
            else RingTeardown();
        }
        InterlockedExchange(&g_prepared, 1);
        g_owners   = 1;
        Unlock();
        return true;
    }

    void Close()
    {
        if (!g_csReady) return;

        // Only the last owner tears down.
        Lock();
        if (!Prepared()) { Unlock(); return; }
        if (g_owners > 1) { g_owners--; Unlock(); return; }
        // Stop producers before the ring is freed.
        InterlockedExchange(&g_prepared, 0);
        g_owners   = 0;
        Unlock();

        // A producer that saw the session armed may still be depositing: its record
        // has to land before the drain's last pass, in a ring that still exists.
        const bool quiet = WaitForProducers(5000);
        // Drain and join OUTSIDE the lock: the drain does not take the CS, and
        // joining it under one Open could also be waiting on is a needless
        // tangle.
        const bool joined = quiet && RingTeardown();   // stop, final drain (may create the file), join, free
        Lock();
        if (!joined) g_ringAbandoned = true;
        // No in-file drop marker: the CSV carries only real rows and loss is reported out of
        // band.
        if (g_file != INVALID_HANDLE_VALUE)
        {
            // A writer still running may use the handle, so it is left open rather than closed under it.
            if (joined) { FlushFileBuffers(g_file); CloseHandle(g_file); }
            g_file = INVALID_HANDLE_VALUE;
        }
        // The path and drop count stay for the report.
        const std::wstring wp = g_path.empty() ? g_planned : g_path;
        const long long    n  = g_rows;
        Unlock();
        if (!quiet)
            core::Log::Warning("trace file: a traced thread was still writing 5s after disarm -- the ring "
                               "and the file are left to it rather than freed, and later sessions write "
                               "synchronously");

        // Where the output went, said at the end as well as the start. The row
        // count is not decoration: a session that traced nothing leaves no file,
        // so the path alone would name something that does not exist.
        if (!wp.empty())
        {
            const std::string p = Narrow(wp);
            char tail[64];
            if (n > 0) _snprintf_s(tail, sizeof tail, _TRUNCATE, "  (%lld row(s))", n);
            else       _snprintf_s(tail, sizeof tail, _TRUNCATE,
                                   "  -- NOT CREATED: nothing was traced");
            core::Log::Info(std::string("trace file (disarmed): ") + p + tail);
        }
    }

    // "Armed", not "file exists" -- see g_prepared.
    bool IsOpen() { return Prepared(); }
    std::wstring Path()
    {
        if (!g_csReady) return std::wstring();
        Lock();
        // the planned name until the file exists
        std::wstring p = g_path.empty() ? g_planned : g_path;
        Unlock();
        return p;
    }
    bool FileExists()
    {
        if (!g_csReady) return false;
        Lock();
        const bool made = !g_path.empty();
        Unlock();
        return made;
    }
    long long RowsWritten() { return g_rows; }
    std::size_t RingCapacityBytes() { return g_ring.CapacityBytes(); }
    long long   RingDrops() { return g_ring.Drops(); }
    long long   RingPauses() { return g_ring.Pauses(); }
    unsigned long long NextSpan() { return static_cast<unsigned long long>(InterlockedIncrement64(&g_span)); }

    namespace
    {
        void WriteCounted(const Row& row)
        {
        // Into the per-thread heap scratch, not the stack. A failed allocation drops the row
        // rather than faulting the hook.
        //
        // The PRODUCER stamps `input` here -- consumed whether or not the
        // row reaches the ring, so a dropped row leaves a HOLE, which is how the
        // file shows where it lost data. `seq` is the writer's, and dense.
        const char* frag = nullptr;
        int n = 0;
        if (Jsonl())
        {
            // Built whole, so a row past kMaxRowBytes is found at the end: its `input` is
            // then a hole, as any lost row's is.
            t_frag.limit = kMaxRowBytes;
            t_frag.Clear();
            char pre[40];
            _snprintf_s(pre, _TRUNCATE, "\"input\":%lld,",
                        static_cast<long long>(InterlockedIncrement64(&g_in)));
            t_frag.Append(pre);
            json::AppendRow(row, t_frag);
            if (t_frag.Over()) return;
            frag = t_frag.Text();
            n = static_cast<int>(t_frag.Len());
        }
        else
        {
        char* buf = FragBuf(24 + FragmentSize(row) + 1);
        if (!buf) return;
        const int il = _snprintf_s(buf, 24, _TRUNCATE, "%lld,",
            static_cast<long long>(InterlockedIncrement64(&g_in)));
        n = il + static_cast<int>(Fragment(row, buf + il));
        frag = buf;
        }   // "input,kind,...,note\r\n"

        if (g_ring.Active())
        {
            g_ring.TryDeposit(frag, n, g_drain.wake);   // lock-free; the drain creates the file
            return;
        }

        // Synchronous: no drain, so the producer is the writer, and the row is
        // on the OS's side of the fence before the call returns.
        Lock();
        // Close may have run since the check above; do not reopen the file.
        if (Prepared() && EnsureFile())
        {
            char seqb[32];
            const int sl = SeqPrefix(static_cast<long long>(++g_seq), seqb, sizeof seqb);
            WriteRaw(seqb, static_cast<DWORD>(sl));
            WriteRaw(frag, static_cast<DWORD>(n));
            g_rows++;
        }
        Unlock();
        }
    }

    void WriteRow(const Row& row)
    {
        if (!g_csReady) return;
        // COUNTED IN BEFORE THE ARMED CHECK: Close clears the flag and then waits
        // for this count, so no producer that saw it set is missed.
        InterlockedIncrement(&g_producers);
        if (Prepared()) WriteCounted(row);   // armed? -- the file may not exist yet
        InterlockedDecrement(&g_producers);
    }

    void ReleaseThreadScratch()
    {
        t_frag.Release();
    }

    void ReleaseHeldByThisThread()
    {
        if (!g_csReady) return;
        while (t_csDepth > 0) { --t_csDepth; LeaveCriticalSection(&g_cs); }
    }

    void FaultWhileLockedForProbe()
    {
        if (!g_csReady) { InitializeCriticalSection(&g_cs); g_csReady = true; }
        Lock();
        volatile int* nowhere = nullptr;
        *nowhere = 1;
    }
}
}   // namespace emit
