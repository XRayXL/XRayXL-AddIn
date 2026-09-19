// DemoBehaviors64.xll -- worksheet functions that show the OTHER columns of a
// trace: string and array returns, an error return, a volatile function that
// re-runs every recalc, and a thread-safe one that Excel spreads across its
// calculation worker threads.
#include "demo_xll_common.h"

using namespace demo;

// A string in, a string out. =ReverseText("hello")  ->  "olleh"
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
ReverseText(const wchar_t* s)
{
    wchar_t buf[512]{};
    size_t len = s ? wcslen(s) : 0; if (len > 500) len = 500;
    // Characters, not UTF-16 units: a surrogate pair (an emoji) stays in order.
    const auto high = [](wchar_t c) { return c >= 0xD800 && c <= 0xDBFF; };
    const auto low  = [](wchar_t c) { return c >= 0xDC00 && c <= 0xDFFF; };
    if (len > 0 && high(s[len - 1])) --len;          // the cap cut a pair in half
    size_t o = 0;
    for (size_t i = len; i > 0; --i)
    {
        if (i >= 2 && low(s[i - 1]) && high(s[i - 2])) { buf[o++] = s[i - 2]; buf[o++] = s[i - 1]; --i; }
        else buf[o++] = s[i - 1];
    }
    buf[o] = 0;
    return RetStr(buf);
}

// Returns a 1 x n array of squares {1, 4, 9, ...}, so a modern Excel SPILLS it.
// One XLL call, many result cells. =MakeSeries(5)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
MakeSeries(double count)
{
    int n = static_cast<int>(count);
    if (n < 1) n = 1; if (n > 256) n = 256;
    double v[256];
    for (int i = 0; i < n; i++) v[i] = static_cast<double>((i + 1) * (i + 1));
    return RetRow(v, n);
}

// May ERROR: returns a/b, or #DIV/0! when b is zero. The error shows in the
// trace's return column. =MightDivide(10, 0)  ->  #DIV/0!
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
MightDivide(double a, double b)
{
    if (b == 0.0) return RetErr(xlerrDiv0);
    return RetNum(a / b);
}

// The same, made safe -- returns 0 instead of erroring. Trace the two side by
// side to see the difference. =SafeDivide(10, 0)  ->  0
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
SafeDivide(double a, double b)
{
    if (b == 0.0) return RetNum(0.0);
    return RetNum(a / b);
}

// THREAD-SAFE ($ in the type string): Excel is free to run it on any of its
// calculation worker threads. Fill a column with =ThreadSafeSquare(...) and the
// trace shows it running under SEVERAL different thread ids.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
ThreadSafeSquare(double x)
{
    return RetNum(x * x);
}

// VOLATILE (! in the type string): recalculates on EVERY calculation, whether or
// not its inputs changed -- so it appears in the trace of every recalc.
// =CallCounter()
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
CallCounter()
{
    static volatile LONG g_count = 0;
    return RetNum(static_cast<double>(InterlockedIncrement(&g_count)));
}

extern "C" __declspec(dllexport) int __stdcall xlAutoOpen()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);

    Register(xDLL, L"ReverseText",      L"QC%");   // wide string in, string out
    Register(xDLL, L"MakeSeries",       L"QB");    // array (spilling) result
    Register(xDLL, L"MightDivide",      L"QBB");   // can return an error
    Register(xDLL, L"SafeDivide",       L"QBB");
    Register(xDLL, L"ThreadSafeSquare", L"QB$");   // $ -> multi-threaded calc
    Register(xDLL, L"CallCounter",      L"Q!");    // ! -> volatile

    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}
