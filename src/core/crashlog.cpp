#include "crashlog.h"
#include "log.h"
#include <windows.h>
#include <psapi.h>
#include <cstddef>

namespace core
{

namespace
{
    wchar_t g_path[MAX_PATH]{};
    LPTOP_LEVEL_EXCEPTION_FILTER g_previous = nullptr;
    // Once per process: installing Filter again hands back Filter as "previous",
    // and a crash would then call itself until the stack ran out.
    volatile LONG g_filterInstalled   = 0;
    volatile LONG g_vectoredInstalled = 0;

    // A FILE CALLED "crash" MUST MEAN A CRASH, so notes are buffered and
    // written only when a fault report is. The last thing noted before a crash
    // still leads the report; a hard kill discards them, which is the right
    // trade -- they are session state, already in the levelled log, not
    // evidence about a death.
    //
    // Fixed storage, no allocation: reachable from the crash path.
    char          g_notes[8192]{};
    volatile LONG g_notesLen  = 0;
    volatile LONG g_notesLost = 0;   // dropped once the buffer filled
    volatile LONG g_flushed   = 0;

    void AppendRaw(const char* text, DWORD len)
    {
        HANDLE h = CreateFileW(g_path, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE,
                               nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;
        DWORD written = 0;
        WriteFile(h, text, len, &written, nullptr);
        CloseHandle(h);
    }

    void AppendText(const char* text)
    {
        DWORD n = 0;
        while (text[n] != 0 && n < 4096) n++;
        AppendRaw(text, n);
    }

    // Hex, written by hand: no CRT formatting at crash time.
    void AppendHex(ULONG64 v)
    {
        char buf[19];
        buf[0] = '0'; buf[1] = 'x';
        for (int i = 0; i < 16; i++)
        {
            int nib = static_cast<int>((v >> ((15 - i) * 4)) & 0xF);
            buf[2 + i] = static_cast<char>(nib < 10 ? ('0' + nib) : ('a' + nib - 10));
        }
        AppendRaw(buf, 18);
    }

    // ---- executable regions we allocated ourselves -------------------------
    struct ExecRegion { ULONG64 base; ULONG64 end; char what[48]; };
    ExecRegion g_regions[16]{};
    volatile LONG g_regionCount = 0;

    wchar_t g_dumpDir[MAX_PATH]{};
    volatile LONG g_reported = 0;      // one-shot: the first such crash is the evidence
    volatile LONG g_dumpEnabled = 0;   // opt-in; see crashlog.h for why

    // SAFE memory probe: ReadProcessMemory returns FALSE instead of faulting,
    // which is what we need inside a vectored handler -- a nested fault here
    // would replace the evidence with a different crash.
    bool SafeRead(ULONG64 addr, void* out, SIZE_T bytes)
    {
        SIZE_T got = 0;
        return ReadProcessMemory(GetCurrentProcess(), reinterpret_cast<LPCVOID>(addr),
                                 out, bytes, &got) && got == bytes;
    }

    // Names an address: a loaded module, one of our own stub pages, or nothing.
    // "nothing" is itself the finding -- it is the signature of the crash this
    // handler exists for.
    void AppendWhere(ULONG64 addr)
    {
        if (addr == 0) { AppendText("<null>"); return; }
        HMODULE mod = nullptr;
        if (GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS
                                 | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                               reinterpret_cast<LPCWSTR>(addr), &mod) && mod != nullptr)
        {
            char name[MAX_PATH]{};
            if (GetModuleBaseNameA(GetCurrentProcess(), mod, name, MAX_PATH) > 0)
                AppendText(name);
            else
                AppendText("<module>");
            AppendText("+");
            AppendHex(addr - reinterpret_cast<ULONG64>(mod));
            return;
        }
        const LONG count = g_regionCount;
        for (LONG i = 0; i < count && i < 16; i++)
        {
            if (addr >= g_regions[i].base && addr < g_regions[i].end)
            {
                AppendText(g_regions[i].what);
                AppendText("+");
                AppendHex(addr - g_regions[i].base);
                return;
            }
        }
        AppendText("<no module>");
    }

    void AppendReg(const char* name, ULONG64 v)
    {
        AppendText("\r\n    ");
        AppendText(name);
        AppendText(" ");
        AppendHex(v);
    }

    void FlushNotes()
    {
        if (InterlockedCompareExchange(&g_flushed, 1, 0) != 0) return;
        const LONG n = InterlockedCompareExchange(&g_notesLen, 0, 0);
        if (n > 0) AppendRaw(g_notes, static_cast<DWORD>(n));
        if (InterlockedCompareExchange(&g_notesLost, 0, 0) > 0)
            AppendText("[some session notes were dropped: buffer full]\r\n");
    }

    // dbghelp is loaded on demand, so nothing links against it and nothing is
    // paid for until the one crash that needs it.
    void WriteDump(EXCEPTION_POINTERS* ep)
    {
        if (g_dumpDir[0] == 0) return;
        if (InterlockedCompareExchange(&g_dumpEnabled, 0, 0) == 0)
        {
            // Say so: the absence of a dump otherwise reads as a second
            // failure.
            AppendText("\r\n  dump: not written -- opt-in; set XRAYXL_CRASHDUMP=1 "
                       "before starting Excel (it contains workbook memory)\r\n");
            return;
        }
        // A DLL load inside a vectored handler, under whatever lock the fault holds:
        // tolerable only because the dump is opt-in and the process is already failing.
        HMODULE dbg = LoadLibraryW(L"dbghelp.dll");
        if (!dbg) { AppendText("\r\n  dump: dbghelp would not load\r\n"); return; }
        typedef BOOL (WINAPI *PFN)(HANDLE, DWORD, HANDLE, int, void*, void*, void*);
        PFN write = reinterpret_cast<PFN>(
            reinterpret_cast<void*>(GetProcAddress(dbg, "MiniDumpWriteDump")));
        if (!write) { AppendText("\r\n  dump: no MiniDumpWriteDump\r\n"); FreeLibrary(dbg); return; }

        wchar_t path[MAX_PATH]{};
        int i = 0;
        for (; g_dumpDir[i] != 0 && i < MAX_PATH - 40; i++) path[i] = g_dumpDir[i];
        const wchar_t* leaf = L"\\XRayXL_exec_fault.dmp";
        for (int j = 0; leaf[j] != 0 && i < MAX_PATH - 1; j++, i++) path[i] = leaf[j];
        path[i] = 0;

        HANDLE f = CreateFileW(path, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS,
                               FILE_ATTRIBUTE_NORMAL, nullptr);
        if (f == INVALID_HANDLE_VALUE) { AppendText("\r\n  dump: could not create file\r\n"); FreeLibrary(dbg); return; }

        // MINIDUMP_EXCEPTION_INFORMATION is declared under pshpack4.h, with the pointer at
        // offset 4. Laid out naturally, dbghelp read a garbage pointer: ERROR_NOACCESS.
#pragma pack(push, 4)
        struct DumpInfo { DWORD tid; EXCEPTION_POINTERS* ep; BOOL client; };
#pragma pack(pop)
        static_assert(sizeof(DumpInfo) == 16 && offsetof(DumpInfo, ep) == 4,
                      "MINIDUMP_EXCEPTION_INFORMATION is 4-byte packed");
        DumpInfo info{ GetCurrentThreadId(), ep, FALSE };
        // Full memory first, then fall back: a full dump of Excel is hundreds
        // of megabytes and the call DOES fail, while a leaner dump still
        // carries the faulting thread's stack.
        const int kFull = 2 | 4 | 0x1000;   // FullMemory|HandleData|ThreadInfo
        const int kLean = 0x1000 | 0x40;    // ThreadInfo|IndirectlyReferencedMemory
        BOOL ok = write(GetCurrentProcess(), GetCurrentProcessId(), f, kFull,
                        &info, nullptr, nullptr);
        const DWORD err = ok ? 0 : GetLastError();
        if (!ok)
        {
            SetFilePointer(f, 0, nullptr, FILE_BEGIN);
            SetEndOfFile(f);
            ok = write(GetCurrentProcess(), GetCurrentProcessId(), f, kLean,
                       &info, nullptr, nullptr);
        }
        if (!ok)
        {
            // Last resort, no exception record. Every thread's stack is still in it.
            SetFilePointer(f, 0, nullptr, FILE_BEGIN);
            SetEndOfFile(f);
            ok = write(GetCurrentProcess(), GetCurrentProcessId(), f, kLean,
                       nullptr, nullptr, nullptr);
        }
        CloseHandle(f);
        FreeLibrary(dbg);
        if (ok) { AppendText("\r\n  dump: written\r\n"); }
        else
        {
            // Empty rather than wrong: CreateFileW has already made the file,
            // and a zero-byte .dmp beside the log reads as evidence.
            DeleteFileW(path);
            AppendText("\r\n  dump: MiniDumpWriteDump failed, file removed -- err ");
            AppendHex(err);
            AppendText("\r\n");
        }
    }

    // The one crash an unhandled filter cannot describe: the instruction
    // pointer is in no mapped module, so execution was transferred to data. An
    // address alone says nothing, so this writes the stack that got there.
    LONG CALLBACK Vectored(EXCEPTION_POINTERS* ep)
    {
        if (!ep || !ep->ExceptionRecord || !ep->ContextRecord) return EXCEPTION_CONTINUE_SEARCH;
        const DWORD code = ep->ExceptionRecord->ExceptionCode;
        if (code != EXCEPTION_ACCESS_VIOLATION &&
            code != EXCEPTION_ILLEGAL_INSTRUCTION &&
            code != EXCEPTION_PRIV_INSTRUCTION)
            return EXCEPTION_CONTINUE_SEARCH;

        const CONTEXT* c = ep->ContextRecord;
        const ULONG64 rip = c->Rip;

        // Narrow by design: our own guarded reads fault inside this module
        // on the path of every VBA statement, and must stay free.
        // RtlPcToFileHeader takes no loader lock, which GetModuleHandleEx would, on
        // every one of those faults, on a calc thread.
        PVOID base = nullptr;
        if (RtlPcToFileHeader(reinterpret_cast<PVOID>(rip), &base) != nullptr)
            return EXCEPTION_CONTINUE_SEARCH;

        if (InterlockedCompareExchange(&g_reported, 1, 0) != 0)
            return EXCEPTION_CONTINUE_SEARCH;      // the first one is the evidence

        FlushNotes();   // only here is it a real fault worth a file
        AppendText("\r\nEXECUTION IN UNMAPPED MEMORY (first chance)\r\n  code ");
        AppendHex(code);
        AppendText("\r\n  rip  ");
        AppendHex(rip);
        AppendText("  ");
        AppendWhere(rip);
        if (code == EXCEPTION_ACCESS_VIOLATION &&
            ep->ExceptionRecord->NumberParameters >= 2)
        {
            AppendText("\r\n  op   ");
            AppendHex(ep->ExceptionRecord->ExceptionInformation[0]);   // 0 r, 1 w, 8 exec
            AppendText("  at ");
            AppendHex(ep->ExceptionRecord->ExceptionInformation[1]);
        }
        AppendText("\r\n  thread ");
        AppendHex(GetCurrentThreadId());

        AppendText("\r\n  registers:");
        AppendReg("rsp", c->Rsp); AppendReg("rbp", c->Rbp); AppendReg("rax", c->Rax);
        AppendReg("rbx", c->Rbx); AppendReg("rcx", c->Rcx); AppendReg("rdx", c->Rdx);
        AppendReg("rsi", c->Rsi); AppendReg("rdi", c->Rdi); AppendReg("r10", c->R10);
        AppendReg("r14", c->R14); AppendReg("r15", c->R15);

        // Who jumped here: a JMP leaves no return address, but the frames that
        // led to the dispatch are still on the stack, so every qword naming
        // code is a candidate caller.
        AppendText("\r\n  stack (qwords from rsp that name code):\r\n");
        for (int i = 0; i < 64; i++)
        {
            ULONG64 v = 0;
            if (!SafeRead(c->Rsp + static_cast<ULONG64>(i) * 8, &v, 8)) break;
            if (v < 0x10000) continue;
            HMODULE m2 = nullptr;
            const bool inModule =
                GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS
                                     | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                                   reinterpret_cast<LPCWSTR>(v), &m2) && m2 != nullptr;
            bool inOurs = false;
            const LONG rcount = g_regionCount;
            for (LONG r = 0; r < rcount && r < 16; r++)
                if (v >= g_regions[r].base && v < g_regions[r].end) inOurs = true;
            if (!inModule && !inOurs) continue;
            AppendText("    [rsp+");
            AppendHex(static_cast<ULONG64>(i) * 8);
            AppendText("] ");
            AppendHex(v);
            AppendText("  ");
            AppendWhere(v);
            AppendText("\r\n");
        }

        WriteDump(ep);
        return EXCEPTION_CONTINUE_SEARCH;          // never change behaviour
    }

    LONG WINAPI Filter(EXCEPTION_POINTERS* ep)
    {
        FlushNotes();   // session notes belong ahead of the report
        AppendText("\r\nUNHANDLED EXCEPTION\r\n  code ");
        AppendHex(ep && ep->ExceptionRecord ? ep->ExceptionRecord->ExceptionCode : 0);
        AppendText("\r\n  at   ");
        ULONG64 addr = 0;
        if (ep && ep->ExceptionRecord)
            addr = reinterpret_cast<ULONG64>(ep->ExceptionRecord->ExceptionAddress);
        AppendHex(addr);

        // Whose code is that?
        AppendText("\r\n  in   ");
        AppendWhere(addr);
        AppendText("\r\n  thread ");
        AppendHex(GetCurrentThreadId());
        AppendText("\r\n");

        return g_previous ? g_previous(ep) : EXCEPTION_CONTINUE_SEARCH;
    }
}

namespace crashlog
{
    // Refuse an over-long path; a truncated one would send the report elsewhere.
    void Install(const std::wstring& path)
    {
        if (path.size() >= MAX_PATH)
        {
            g_path[0] = 0;
            Log::Warning("crash report: the path is longer than MAX_PATH and has been "
                         "refused -- a crash in this session will not write one. Set "
                         "XRAYXL_OUTPUT_DIR to something shorter.");
        }
        else
        {
            for (size_t i = 0; i < path.size(); i++) g_path[i] = path[i];
            g_path[path.size()] = 0;
        }
        if (InterlockedCompareExchange(&g_filterInstalled, 1, 0) == 0)
            g_previous = SetUnhandledExceptionFilter(Filter);
    }

    void InstallVectored(const std::wstring& dumpDir)
    {
        // Same; the dump path needs 40 characters for the file name.
        if (dumpDir.size() >= MAX_PATH - 40)
        {
            g_dumpDir[0] = 0;
            Log::Warning("crash dump: the output directory is too long to build a dump "
                         "path in and has been refused -- the text report is unaffected.");
        }
        else
        {
            for (size_t i = 0; i < dumpDir.size(); i++) g_dumpDir[i] = dumpDir[i];
            g_dumpDir[dumpDir.size()] = 0;
        }

        // Read ONCE: the handler goes in at load, long before anything can be
        // armed, which is why this is an environment variable.
        wchar_t v[8]{};
        const DWORD got = GetEnvironmentVariableW(L"XRAYXL_CRASHDUMP", v, 8);
        const bool on = (got > 0 && got < 8) &&
                        (v[0] == L'1' || v[0] == L'y' || v[0] == L'Y' ||
                         v[0] == L't' || v[0] == L'T');
        InterlockedExchange(&g_dumpEnabled, on ? 1 : 0);

        if (InterlockedCompareExchange(&g_vectoredInstalled, 1, 0) == 0)
            AddVectoredExceptionHandler(1, Vectored);   // 1 = first in the chain

        // Which mode was in force is the first question asked of a crash that
        // produced no dump. Buffered like every other note, so a session that
        // does not fault still leaves no file.
        Note(on ? "crash capture: text always; minidump ENABLED (XRAYXL_CRASHDUMP)"
                : "crash capture: text always; minidump off (set XRAYXL_CRASHDUMP=1 to enable)");
    }

    void SetDumpEnabled(bool on) { InterlockedExchange(&g_dumpEnabled, on ? 1 : 0); }

    void NoteExecRegion(const void* base, std::size_t bytes, const char* what)
    {
        const LONG i = InterlockedIncrement(&g_regionCount) - 1;
        if (i < 0 || i >= 16) return;
        g_regions[i].base = reinterpret_cast<ULONG64>(base);
        g_regions[i].end  = g_regions[i].base + bytes;
        int j = 0;
        for (; what[j] != 0 && j < 47; j++) g_regions[i].what[j] = what[j];
        g_regions[i].what[j] = 0;
    }

    void Note(const char* text)
    {
        int n = 0;
        while (text[n] != 0 && n < 512) ++n;
        const LONG at = InterlockedExchangeAdd(&g_notesLen, n + 2);
        if (at + n + 2 >= static_cast<LONG>(sizeof(g_notes)))
        {
            InterlockedExchangeAdd(&g_notesLen, -(n + 2));   // give the space back
            InterlockedIncrement(&g_notesLost);
            return;
        }
        for (int i = 0; i < n; ++i) g_notes[at + i] = text[i];
        g_notes[at + n]     = 0x0D;
        g_notes[at + n + 1] = 0x0A;
    }
}
}   // namespace core
