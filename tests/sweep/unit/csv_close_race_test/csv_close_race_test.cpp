// Unit test for emit::csv::Close (emit/csv.cpp): the output ring must outlive every producer
// that saw the session armed. Writer threads deposit rows nonstop while the session is opened
// and closed in a loop; a 16 KB PAUSE ring keeps them inside TryDeposit waiting for room.
//
// It guards against a close that faults or hangs under writers. It cannot see a late deposit
// into freed heap, which stays mapped. Needs no Excel; the log and the output root are stubbed.
//
// Built by XRayXL.sln into build\x64\Release\unit\.

#include "emit/csv.h"

#include <windows.h>
#include <cstdio>
#include <string>

namespace
{
    std::wstring g_dir;
    volatile LONG   g_stop = 0;
    volatile LONG64 g_written = 0;

    DWORD WINAPI Writer(LPVOID)
    {
        emit::csv::Row r;
        r.kind = "entry"; r.source = "XLL"; r.function = "TxHammered"; r.args = "a1:B=1";
        while (InterlockedCompareExchange(&g_stop, 0, 0) == 0)
        {
            emit::csv::WriteRow(r);
            InterlockedIncrement64(&g_written);
        }
        return 0;
    }

    // A fault on a writer thread would end the process silently; say what it was.
    LONG WINAPI OnFault(EXCEPTION_POINTERS* ep)
    {
        printf("[FAIL] a writer faulted (code 0x%08lX) -- the ring was freed under it\n\nFAILED\n",
               ep && ep->ExceptionRecord ? ep->ExceptionRecord->ExceptionCode : 0UL);
        fflush(stdout);
        TerminateProcess(GetCurrentProcess(), 1);
        return EXCEPTION_EXECUTE_HANDLER;
    }

    void ClearDir()
    {
        WIN32_FIND_DATAW fd;
        HANDLE h = FindFirstFileW((g_dir + L"\\*.csv").c_str(), &fd);
        if (h == INVALID_HANDLE_VALUE) return;
        do { DeleteFileW((g_dir + L"\\" + fd.cFileName).c_str()); } while (FindNextFileW(h, &fd));
        FindClose(h);
    }
}

namespace core
{
    namespace Log
    {
        void Info(const std::string&) {}
        void Warning(const std::string&) {}
        void Note(const std::string&) {}
    }
    std::wstring EnsureAppSubdir(const wchar_t*) { return g_dir; }
}

int main()
{
    wchar_t tmp[MAX_PATH];
    GetTempPathW(MAX_PATH, tmp);
    g_dir = std::wstring(tmp) + L"xray_csv_close_race";
    CreateDirectoryW(g_dir.c_str(), nullptr);
    ClearDir();
    SetUnhandledExceptionFilter(OnFault);

    constexpr int kWriters = 4;
    constexpr int kCycles  = 200;
    HANDLE threads[kWriters];
    for (int i = 0; i < kWriters; ++i) threads[i] = CreateThread(nullptr, 0, Writer, nullptr, 0, nullptr);

    for (int c = 0; c < kCycles; ++c)
    {
        emit::csv::Open(16 * 1024, true, core::modes::Format::Csv, false);
        Sleep(1);
        emit::csv::Close();
    }

    // A synchronous session after a dropping ring reports no drops: the counts survive Close
    // for the summary and must not reach the next session's Disarm result.
    emit::csv::Open(16 * 1024, false, core::modes::Format::Csv, false);
    for (int i = 0; i < 5000 && emit::csv::RingDrops() == 0; ++i) Sleep(1);
    const long long ringDrops = emit::csv::RingDrops();
    emit::csv::Close();
    emit::csv::Open(0, false, core::modes::Format::Csv, false);
    const long long syncDrops = emit::csv::RingDrops();
    emit::csv::Close();

    InterlockedExchange(&g_stop, 1);
    WaitForMultipleObjects(kWriters, threads, TRUE, 10000);
    ClearDir();

    printf("  %d open/close cycles under %d writers, %lld rows offered\n",
           kCycles, kWriters, static_cast<long long>(g_written));
    printf("[PASS] no writer faulted while the session closed under it\n");
    const bool countsOk = ringDrops > 0 && syncDrops == 0;
    printf("[%s] a synchronous session after a 16 KB DROP ring: ring dropped %lld, synchronous reports %lld (must be > 0, then 0)\n",
           countsOk ? "PASS" : "FAIL", ringDrops, syncDrops);
    printf("\n%s\n", countsOk ? "ALL PASS" : "FAILED");
    return countsOk ? 0 : 1;
}
