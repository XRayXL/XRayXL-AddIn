#include "snapshot.h"

#include <windows.h>
#include <tlhelp32.h>
#include <winver.h>

#include <algorithm>
#include <cstdarg>
#include <cwchar>

// A column that cannot be read is left empty rather than failing the table: a module whose file
// is gone still deserves its name, base and size.

namespace diag
{
namespace
{
    std::wstring Fmt(const wchar_t* form, ...)
    {
        wchar_t buf[512];
        va_list args;
        va_start(args, form);
        const int n = _vsnwprintf_s(buf, _TRUNCATE, form, args);
        va_end(args);
        return std::wstring(buf, n > 0 ? n : 0);
    }

    // Thousands-separated, so a byte count can be compared by eye between two machines.
    std::wstring Grouped(unsigned long long v)
    {
        wchar_t plain[32];
        _snwprintf_s(plain, _TRUNCATE, L"%llu", v);
        std::wstring out;
        const int len = static_cast<int>(wcslen(plain));
        for (int i = 0; i < len; ++i)
        {
            if (i > 0 && (len - i) % 3 == 0) out += L',';
            out += plain[i];
        }
        return out;
    }

    std::wstring KB(unsigned long long bytes) { return Grouped(bytes / 1024) + L" KB"; }

    std::wstring LocalStamp(const FILETIME& ft)
    {
        FILETIME local{};
        SYSTEMTIME st{};
        if (!FileTimeToLocalFileTime(&ft, &local) || !FileTimeToSystemTime(&local, &st)) return L"";
        return Fmt(L"%04u-%02u-%02u %02u:%02u:%02u",
                   st.wYear, st.wMonth, st.wDay, st.wHour, st.wMinute, st.wSecond);
    }

    std::wstring Duration(unsigned long long hundredNs)
    {
        const unsigned long long secs = hundredNs / 10000000ull;
        return Fmt(L"%llu:%02llu:%02llu", secs / 3600, (secs / 60) % 60, secs % 60);
    }

    // ---- version resources ---------------------------------------------------------------

    struct FileFacts
    {
        std::wstring version;
        std::wstring modified;
        std::wstring size;
    };

    // Microsoft's system files append their build lab to FileVersion -- "10.0.26100.9278
    // (WinBuild.160101.0800)". The tag is the same on every machine, so it is noise here.
    std::wstring WithoutBuildLab(std::wstring v)
    {
        const size_t at = v.find(L" (");
        if (at != std::wstring::npos) v.resize(at);
        while (!v.empty() && v.back() == L' ') v.pop_back();
        return v;
    }

    // The string block Windows Explorer shows, in the file's own language before English.
    std::wstring StringValue(const void* block, DWORD langCodepage, const wchar_t* name)
    {
        wchar_t path[128];
        _snwprintf_s(path, _TRUNCATE, L"\\StringFileInfo\\%04x%04x\\%s",
                     LOWORD(langCodepage), HIWORD(langCodepage), name);
        void* text = nullptr;
        UINT  chars = 0;
        if (!VerQueryValueW(block, path, &text, &chars) || !text || chars == 0) return L"";
        std::wstring out(static_cast<const wchar_t*>(text), chars);
        while (!out.empty() && out.back() == L'\0') out.pop_back();
        return out;
    }

    FileFacts ReadFile(const std::wstring& path)
    {
        FileFacts f;

        WIN32_FILE_ATTRIBUTE_DATA a{};
        if (GetFileAttributesExW(path.c_str(), GetFileExInfoStandard, &a))
        {
            f.modified = LocalStamp(a.ftLastWriteTime);
            f.size = Grouped((static_cast<unsigned long long>(a.nFileSizeHigh) << 32) | a.nFileSizeLow);
        }
        DWORD ignored = 0;
        const DWORD bytes = GetFileVersionInfoSizeW(path.c_str(), &ignored);
        if (!bytes) return f;
        std::vector<unsigned char> block(bytes);
        if (!GetFileVersionInfoW(path.c_str(), 0, bytes, block.data())) return f;

        void* translation = nullptr;
        UINT  tbytes = 0;
        if (VerQueryValueW(block.data(), L"\\VarFileInfo\\Translation", &translation, &tbytes) && translation && tbytes >= 4)
        {
            const DWORD* langs = static_cast<const DWORD*>(translation);
            const unsigned count = tbytes / 4;
            for (unsigned i = 0; i < count && f.version.empty(); ++i)
                f.version = WithoutBuildLab(StringValue(block.data(), langs[i], L"FileVersion"));
        }
        // The binary FILEVERSION is the truth when the string block is missing or decorative.
        void* fixed = nullptr;
        UINT  fbytes = 0;
        if (f.version.empty() && VerQueryValueW(block.data(), L"\\", &fixed, &fbytes) && fixed && fbytes >= sizeof(VS_FIXEDFILEINFO))
        {
            const VS_FIXEDFILEINFO* fi = static_cast<const VS_FIXEDFILEINFO*>(fixed);
            f.version = Fmt(L"%u.%u.%u.%u", HIWORD(fi->dwFileVersionMS), LOWORD(fi->dwFileVersionMS),
                            HIWORD(fi->dwFileVersionLS), LOWORD(fi->dwFileVersionLS));
        }
        return f;
    }

    // ---- the counters user32 and kernel32 keep about us ------------------------------------

    unsigned GuiResources(DWORD flag)
    {
        typedef DWORD (WINAPI* Fn)(HANDLE, DWORD);
        static const auto fn = reinterpret_cast<Fn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetGuiResources")));
        return fn ? fn(GetCurrentProcess(), flag) : 0;
    }

    struct MemoryCounters
    {
        SIZE_T workingSet = 0, peakWorkingSet = 0, privateBytes = 0, peakPrivate = 0;
        DWORD  pageFaults = 0;
        bool   ok = false;
    };

    MemoryCounters ReadMemory()
    {
        // PROCESS_MEMORY_COUNTERS_EX, declared here so psapi.h is not needed.
        struct Counters
        {
            DWORD  cb;
            DWORD  PageFaultCount;
            SIZE_T PeakWorkingSetSize, WorkingSetSize;
            SIZE_T QuotaPeakPagedPoolUsage, QuotaPagedPoolUsage;
            SIZE_T QuotaPeakNonPagedPoolUsage, QuotaNonPagedPoolUsage;
            SIZE_T PagefileUsage, PeakPagefileUsage;
            SIZE_T PrivateUsage;
        };
        typedef BOOL (WINAPI* Fn)(HANDLE, Counters*, DWORD);
        static const auto fn = reinterpret_cast<Fn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"kernel32.dll"), "K32GetProcessMemoryInfo")));

        MemoryCounters out;
        Counters c{};
        c.cb = sizeof(c);
        if (fn && fn(GetCurrentProcess(), &c, sizeof(c)))
        {
            out.workingSet = c.WorkingSetSize;
            out.peakWorkingSet = c.PeakWorkingSetSize;
            out.privateBytes = c.PrivateUsage;
            out.peakPrivate = c.PeakPagefileUsage;
            out.pageFaults = c.PageFaultCount;
            out.ok = true;
        }
        return out;
    }

    unsigned ThreadCount()
    {
        HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0);
        if (snap == INVALID_HANDLE_VALUE) return 0;
        const DWORD mine = GetCurrentProcessId();
        unsigned n = 0;
        THREADENTRY32 t{};
        t.dwSize = sizeof(t);
        for (BOOL more = Thread32First(snap, &t); more; more = Thread32Next(snap, &t))
            if (t.th32OwnerProcessID == mine) ++n;
        CloseHandle(snap);
        return n;
    }

    // The real build, not the manifest's: GetVersionEx lies to an unmanifested caller.
    std::wstring OsVersion()
    {
        typedef LONG (WINAPI* Fn)(PRTL_OSVERSIONINFOW);
        static const auto fn = reinterpret_cast<Fn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"ntdll.dll"), "RtlGetVersion")));
        RTL_OSVERSIONINFOW v{};
        v.dwOSVersionInfoSize = sizeof(v);
        if (!fn || fn(&v) != 0) return L"";
        return Fmt(L"%lu.%lu build %lu", v.dwMajorVersion, v.dwMinorVersion, v.dwBuildNumber);
    }

    void Add(Table& t, const wchar_t* name, const std::wstring& value)
    {
        t.rows.push_back({ name, value });
    }
}

Table Modules()
{
    Table t;
    t.columns = { { L"Module", 162, false }, { L"Version", 119, false },
                  { L"Date Modified", 120, false }, { L"Size", 73, true },
                  { L"Path", 420, false } };

    HANDLE snap = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId());
    if (snap == INVALID_HANDLE_VALUE) return t;

    MODULEENTRY32W m{};
    m.dwSize = sizeof(m);
    for (BOOL more = Module32FirstW(snap, &m); more; more = Module32NextW(snap, &m))
    {
        const FileFacts f = ReadFile(m.szExePath);
        t.rows.push_back({ m.szModule, f.version, f.modified, f.size, m.szExePath });
    }
    CloseHandle(snap);

    // The snapshot's own order is load order; by name is what two machines can be compared in.
    // The process image stays first, as the thing everything else was loaded into.
    if (t.rows.size() > 1)
        std::sort(t.rows.begin() + 1, t.rows.end(),
                  [](const std::vector<std::wstring>& a, const std::vector<std::wstring>& b)
                  { return _wcsicmp(a[0].c_str(), b[0].c_str()) < 0; });
    return t;
}

Table Environment()
{
    Table t;
    t.columns = { { L"Variable", 230, false }, { L"Value", 430, false } };

    wchar_t* block = GetEnvironmentStringsW();
    if (!block) return t;
    for (const wchar_t* p = block; *p; p += wcslen(p) + 1)
    {
        const std::wstring entry(p);
        // The drive-letter entries Windows hides start with '='; they are current directories.
        const size_t eq = entry.find(L'=', 1);
        if (eq == std::wstring::npos) continue;
        t.rows.push_back({ entry.substr(0, eq), entry.substr(eq + 1) });
    }
    FreeEnvironmentStringsW(block);

    std::sort(t.rows.begin(), t.rows.end(),
              [](const std::vector<std::wstring>& a, const std::vector<std::wstring>& b)
              { return _wcsicmp(a[0].c_str(), b[0].c_str()) < 0; });
    return t;
}

Table Process(const Session& session)
{
    Table t;
    t.columns = { { L"Metric", 230, false }, { L"Value", 430, false } };

    wchar_t image[MAX_PATH] = {};
    GetModuleFileNameW(nullptr, image, MAX_PATH);

    Add(t, L"Process ID", Fmt(L"%lu", GetCurrentProcessId()));
    Add(t, L"Process image", image);
    Add(t, L"Bitness", Fmt(L"%d-bit", static_cast<int>(sizeof(void*) * 8)));

    FILETIME created{}, exited{}, kernel{}, user{};
    if (GetProcessTimes(GetCurrentProcess(), &created, &exited, &kernel, &user))
    {
        ULARGE_INTEGER k{}, u{};
        k.LowPart = kernel.dwLowDateTime; k.HighPart = kernel.dwHighDateTime;
        u.LowPart = user.dwLowDateTime;   u.HighPart = user.dwHighDateTime;
        Add(t, L"Started", LocalStamp(created));
        Add(t, L"CPU time (user)", Duration(u.QuadPart));
        Add(t, L"CPU time (kernel)", Duration(k.QuadPart));
    }

    const MemoryCounters mem = ReadMemory();
    if (mem.ok)
    {
        Add(t, L"Working set", KB(mem.workingSet));
        Add(t, L"Peak working set", KB(mem.peakWorkingSet));
        Add(t, L"Private bytes", KB(mem.privateBytes));
        Add(t, L"Peak private bytes", KB(mem.peakPrivate));
        Add(t, L"Page faults", Grouped(mem.pageFaults));
    }

    DWORD handles = 0;
    if (GetProcessHandleCount(GetCurrentProcess(), &handles)) Add(t, L"Handles", Grouped(handles));
    Add(t, L"GDI objects", Grouped(GuiResources(0 /*GR_GDIOBJECTS*/)));
    Add(t, L"GDI objects (peak)", Grouped(GuiResources(2 /*GR_GDIOBJECTS_PEAK*/)));
    Add(t, L"USER objects", Grouped(GuiResources(1 /*GR_USEROBJECTS*/)));
    Add(t, L"USER objects (peak)", Grouped(GuiResources(4 /*GR_USEROBJECTS_PEAK*/)));
    Add(t, L"Threads", Grouped(ThreadCount()));

    MEMORYSTATUSEX ms{};
    ms.dwLength = sizeof(ms);
    if (GlobalMemoryStatusEx(&ms))
    {
        Add(t, L"Machine memory", KB(ms.ullTotalPhys));
        Add(t, L"Machine memory free", KB(ms.ullAvailPhys));
        Add(t, L"Address space in use", KB(ms.ullTotalVirtual - ms.ullAvailVirtual));
    }

    SYSTEM_INFO si{};
    GetNativeSystemInfo(&si);
    Add(t, L"Processors", Fmt(L"%lu", si.dwNumberOfProcessors));
    Add(t, L"Windows", OsVersion());
    wchar_t machine[MAX_COMPUTERNAME_LENGTH + 1] = {};
    DWORD machineLen = MAX_COMPUTERNAME_LENGTH + 1;
    if (GetComputerNameW(machine, &machineLen)) Add(t, L"Computer", machine);

    wchar_t xrayVersion[32] = {};
    MultiByteToWideChar(CP_UTF8, 0, session.version, -1, xrayVersion, 32);
    Add(t, L"XRayXL version", xrayVersion);
    Add(t, L"XRayXL armed", session.armed ? L"Yes" : L"No");
    Add(t, L"Trace file", session.traceFile.empty() ? L"(named when tracing is armed)" : session.traceFile);
    Add(t, L"Log file", session.logFile);
    return t;
}
}
