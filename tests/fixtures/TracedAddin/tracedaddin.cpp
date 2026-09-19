// TracedAddin: a deliberately ordinary XLL for the tracing tests to watch. Not built from the
// product's sources, so when a trace of it is wrong the fault is the tracer's.
//
// It covers the type codes. Every function's answer depends only on its inputs, so a test can
// assert what the trace says about a call. Naming: Tx<Code> takes that code; TxRet<Code>
// returns it.

#include <windows.h>
#include <cwchar>
#include <cstdio>
#include <cmath>
#include "plain_xll.h"

using plainxll::PascalStr;
using plainxll::RetNum;
using plainxll::RetStr;
using plainxll::RetErr;
using plainxll::RetOwnedStr;
using plainxll::t_ret;

namespace
{
    void Log(const char* what, int detail)
    {
        wchar_t path[MAX_PATH];
        DWORD n = GetTempPathW(MAX_PATH, path);
        if (n == 0) return;
        wcscat_s(path, L"tracedaddin.log");
        HANDLE h = CreateFileW(path, FILE_APPEND_DATA, FILE_SHARE_READ | FILE_SHARE_WRITE,
                               nullptr, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (h == INVALID_HANDLE_VALUE) return;
        char line[256];
        int len = _snprintf_s(line, _TRUNCATE, "%s = %d\r\n", what, detail);
        DWORD written = 0;
        WriteFile(h, line, static_cast<DWORD>(len), &written, nullptr);
        CloseHandle(h);
    }

    // Logs each registration's result, so a #NAME? in a test has a cause on disk.
    void Register(XLOPER12& xDLL, const wchar_t* name, const wchar_t* typeText,
                  const wchar_t* displayName = nullptr)
    {
        const int rc = plainxll::Register(xDLL, name, typeText, displayName);
        char buf[160];
        _snprintf_s(buf, _TRUNCATE, "register %S (%S) rc", name, typeText);
        Log(buf, rc);
    }

    // A command: macro_type 2, the sixth xlfRegister argument.
    void RegisterCommand(XLOPER12& xDLL, const wchar_t* name)
    {
        PascalStr proc; proc.Set(name);
        PascalStr type; type.Set(L"J");
        PascalStr func; func.Set(name);
        PascalStr args; args.Set(L"");
        XLOPER12 macroType; ZeroMemory(&macroType, sizeof(macroType));
        macroType.xltype = xltypeInt; macroType.val.w = 2;
        XLOPER12 res; ZeroMemory(&res, sizeof(res));
        const int rc = Excel12(xlfRegister, &res, 6, &xDLL, &proc.oper, &type.oper,
                               &func.oper, &args.oper, &macroType);
        char buf[160];
        _snprintf_s(buf, _TRUNCATE, "register %S (J, command) rc", name);
        Log(buf, rc);
        Excel12(xlFree, nullptr, 1, &res);
    }

    HMODULE ThisModule()
    {
        HMODULE h = nullptr;
        GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           reinterpret_cast<LPCWSTR>(&ThisModule), &h);
        return h;
    }

    // Per-thread, because an XLL function may run on any calculation worker.
    __declspec(thread) XCHAR    t_wideCounted[256];
    __declspec(thread) char     t_asciiCounted[256];
    __declspec(thread) char     t_asciiZ[256];
    __declspec(thread) wchar_t  t_wideZ[256];
    __declspec(thread) double   t_double;
    __declspec(thread) short    t_short;
    __declspec(thread) int      t_int;

    LPXLOPER12 RetBool(bool b)
    {
        t_ret.xltype = xltypeBool; t_ret.val.xbool = b ? 1 : 0; return &t_ret;
    }
}

// ============================================================================
// ARGUMENT SHAPES -- one exported function per type code.
// Each answer is a pure function of the inputs so a test can assert the value.
// ============================================================================

// B: double, in XMM by position. Two of them, so ORDER is observable.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxB(double a, double b)
{ return RetNum(a * 10.0 + b); }

// Five doubles: the fifth is the first STACK argument, which is the case a
// register-only decoder gets wrong.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxStackArgs(
    double a, double b, double c, double d, double e)
{ return RetNum(a * 10000 + b * 1000 + c * 100 + d * 10 + e); }

// Registered QBB at load and QBBBBB again while the tracer is armed (TxRegisterReshape).
// Excel keeps one registration, so after the second one it calls this to the wider shape.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxTwoShapes(
    double a, double b, double c, double d, double e)
{ return RetNum(a + b + c + d + e); }

// A: boolean as a short.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxA(short flag)
{ return RetNum(flag ? 1 : 0); }

// Exported, and never registered by this XLL, so a test can register it under a different
// display name: the Excel-DNA shape, where the exports are f0, f1, f2. Re-registering an
// already registered function does not reproduce it, because Excel keeps one registration per
// (module, procedure) and adds the second name to the same id.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxUnregistered(short flag)
{ return RetNum(flag ? 4242 : -1); }

// I / J: signed short / signed int32.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxI(short v) { return RetNum(v); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxJ(int v)   { return RetNum(v); }

// C: null-terminated ASCII. D: byte-counted ASCII.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxC(const char* s)
{ return RetNum(s ? static_cast<double>(strlen(s)) : -1.0); }

extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxD(const unsigned char* s)
{ return RetNum(s ? static_cast<double>(s[0]) : -1.0); }

// C%: null-terminated UTF-16. D%: WCHAR-counted UTF-16.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxCw(const wchar_t* s)
{ return RetNum(s ? static_cast<double>(wcslen(s)) : -1.0); }

extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxDw(const XCHAR* s)
{ return RetNum(s ? static_cast<double>(s[0]) : -1.0); }

// E: pointer to double. L/M: pointer to short. N: pointer to int32.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxE(double* p) { return RetNum(p ? *p : -1); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxM(short*  p) { return RetNum(p ? *p : -1); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxN(int*    p) { return RetNum(p ? *p : -1); }

// K: FP (2-byte dimensions). K%: FP12 (4-byte dimensions). The pair that
// cross-decodes to "zero cells" rather than crashing.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxK(FP* a)
{
    if (!a) return RetNum(-1);
    double t = 0; const long long n = static_cast<long long>(a->rows) * a->columns;
    for (long long i = 0; i < n; i++) t += a->array[i];
    return RetNum(t);
}
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxKw(FP12* a)
{
    if (!a) return RetNum(-1);
    double t = 0; const long long n = static_cast<long long>(a->rows) * a->columns;
    for (long long i = 0; i < n; i++) t += a->array[i];
    return RetNum(t);
}

// O%: ONE type code, THREE ABI slots -- rows*, columns*, array. The shape a
// per-character parser miscounts most badly.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxOw(int* rows, int* cols, double* arr)
{
    if (!rows || !cols || !arr) return RetNum(-1);
    double t = 0; const long long n = static_cast<long long>(*rows) * (*cols);
    for (long long i = 0; i < n; i++) t += arr[i];
    return RetNum(t);
}

// P: XLOPER (24 bytes, xltype at +16). Q: XLOPER12 (32 bytes, xltype at +24).
// Two different structs: reading Q as an XLOPER finds the type at the wrong offset.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxP(LPXLOPER v)
{
    if (!v) return RetNum(0);
    const WORD t = v->xltype & 0x0FFF;
    if (t & xltypeNum)  return RetNum(1);
    if (t & xltypeStr)  return RetNum(2);
    if (t & xltypeBool) return RetNum(3);
    return RetNum(9);
}

// Any value at all, reported as a code, so a test can confirm the ARGUMENT
// TYPE survived: 1 num, 2 str, 3 bool, 4 error, 5 array, 6 missing.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxQ(LPXLOPER12 v)
{
    if (!v) return RetNum(0);
    const int t = v->xltype & 0x0FFF;
    if (t & xltypeNum)     return RetNum(1);
    if (t & xltypeStr)     return RetNum(2);
    if (t & xltypeBool)    return RetNum(3);
    if (t & xltypeErr)     return RetNum(4);
    if (t & xltypeMulti)   return RetNum(5);
    if (t & xltypeMissing) return RetNum(6);
    return RetNum(9);
}

// A mixed signature: the shape where a
// per-character walk reports five arguments instead of three.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxMixed(
    const XCHAR* name, FP12* values, double* factor)
{
    double t = 0;
    if (values) { const long long n = static_cast<long long>(values->rows) * values->columns;
                  for (long long i = 0; i < n; i++) t += values->array[i]; }
    return RetNum(t * (factor ? *factor : 1.0) + (name ? name[0] : 0));
}

// An omitted argument arrives as xltypeMissing rather than silently as zero.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxOptional(LPXLOPER12 a, LPXLOPER12 b)
{
    const double av = (a && (a->xltype & xltypeNum)) ? a->val.num : 0.0;
    const bool omitted = (b == nullptr) || ((b->xltype & (xltypeMissing | xltypeNil)) != 0);
    return RetNum(omitted ? av : av + 1000.0);
}

// Exported through a jump table (jumptable.asm), not directly.
extern "C" LPXLOPER12 __stdcall TxJumpAddImpl(double a, double b)
{ return RetNum(a + b + 0.5); }

extern "C" LPXLOPER12 __stdcall TxJumpStrImpl(const XCHAR* s)
{ return RetNum(s ? static_cast<double>(s[0]) * 100.0 : -1.0); }

// ---- nesting: an add-in function that calls another THROUGH Excel ------------
//
// Cell nesting like =TxB(TxB(1,2),3) is NOT nesting at the ABI level: Excel
// evaluates the inner call, then the outer, one after the other. To get our
// hook for one function genuinely on the stack while the hook for another
// fires, the add-in has to re-enter Excel. xlUDF does that.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxCallsBack(double x)
{
    PascalStr fn; fn.Set(L"TxB");
    XLOPER12 a; ZeroMemory(&a, sizeof(a)); a.xltype = xltypeNum; a.val.num = x;
    XLOPER12 b; ZeroMemory(&b, sizeof(b)); b.xltype = xltypeNum; b.val.num = 1;
    XLOPER12 res; ZeroMemory(&res, sizeof(res));
    const int rc = Excel12(xlUDF, &res, 3, &fn.oper, &a, &b);
    double v = -1;
    if (rc == xlretSuccess && (res.xltype & xltypeNum)) v = res.val.num;
    Excel12(xlFree, nullptr, 1, &res);
    return RetNum(v + 0.25);
}

// Two levels of re-entry, so the depth bookkeeping is pushed further than one.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxCallsBack2(double x)
{
    PascalStr fn; fn.Set(L"TxCallsBack");
    XLOPER12 a; ZeroMemory(&a, sizeof(a)); a.xltype = xltypeNum; a.val.num = x;
    XLOPER12 res; ZeroMemory(&res, sizeof(res));
    const int rc = Excel12(xlUDF, &res, 2, &fn.oper, &a);
    double v = -1;
    if (rc == xlretSuccess && (res.xltype & xltypeNum)) v = res.val.num;
    Excel12(xlFree, nullptr, 1, &res);
    return RetNum(v * 2.0);
}

// ---- high arity: twelve doubles, so EIGHT of them are on the stack -----------
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxMany(
    double a1, double a2, double a3, double a4, double a5, double a6,
    double a7, double a8, double a9, double a10, double a11, double a12)
{
    // Position-weighted, so a copy loop that shifts or truncates is visible in
    // the answer rather than merely in the trace.
    return RetNum(a1*1 + a2*2 + a3*3 + a4*4 + a5*5 + a6*6 +
                  a7*7 + a8*8 + a9*9 + a10*10 + a11*11 + a12*12);
}

// ---- a COMMAND (macro), not a worksheet function -----------------------------
// Registered with macro_type 2. Excel calls it through Application.Run, with no
// calling cell at all -- which the trace must report as no cell rather than a
// confident A1.
static volatile LONG64 g_macroCalls = 0;
extern "C" __declspec(dllexport) int __stdcall TxMacro()
{
    InterlockedIncrement64(&g_macroCalls);
    return 1;
}

// How many times the macro actually ran, readable from a formula so a test can
// check the macro fired without trusting the tracer to say so.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxMacroCalls()
{
    return RetNum(static_cast<double>(g_macroCalls));
}

// No arguments and a constant result: typetext must be empty and argcount 0.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxNoArgs()
{
    return RetNum(42);
}

// A function that raises. If the exception unwinds through our thunk, the thunk's unwind data
// (xllthunk.asm) has to be right or the process dies. The add-in catches its own exception;
// what is under test is that our frame survives the unwind.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRaises(double x)
{
    __try
    {
        if (x > 0)
        {
            volatile int* p = nullptr;
            *p = 1;                     // deliberate access violation
        }
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return RetNum(-999);
    }
    return RetNum(x);
}

// A divide that raises an integer exception rather than an access violation,
// so both SEH shapes cross the thunk.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxDivZero(double x)
{
    __try
    {
        volatile int z = 0;
        volatile int r = static_cast<int>(x) / z;
        return RetNum(r);
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        return RetNum(-888);
    }
}

// Deliberately slow, so a stress sheet actually engages several calc threads
// at once rather than finishing before Excel bothers to parallelise.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxSlow(double x)
{
    volatile double acc = 0;
    for (int i = 0; i < 20000; i++) acc += i * 0.000001;
    return RetNum(x + acc - acc);
}

// ============================================================================
// RETURN SHAPES -- one exported function per return code.
// ============================================================================

extern "C" __declspec(dllexport) double __stdcall TxRetB(double x) { return x * 3.0; }

extern "C" __declspec(dllexport) short  __stdcall TxRetI(double x) { return static_cast<short>(x) + 1; }
extern "C" __declspec(dllexport) int    __stdcall TxRetJ(double x) { return static_cast<int>(x) + 2; }

extern "C" __declspec(dllexport) const char* __stdcall TxRetC(double x)
{ _snprintf_s(t_asciiZ, sizeof(t_asciiZ), _TRUNCATE, "c:%.0f", x); return t_asciiZ; }

extern "C" __declspec(dllexport) const unsigned char* __stdcall TxRetD(double x)
{
    char tmp[64]; int n = _snprintf_s(tmp, _TRUNCATE, "d:%.0f", x);
    t_asciiCounted[0] = static_cast<char>(n);
    memcpy(t_asciiCounted + 1, tmp, n);
    return reinterpret_cast<const unsigned char*>(t_asciiCounted);
}

extern "C" __declspec(dllexport) const wchar_t* __stdcall TxRetCw(double x)
{ _snwprintf_s(t_wideZ, 256, _TRUNCATE, L"cw:%.0f", x); return t_wideZ; }

extern "C" __declspec(dllexport) const XCHAR* __stdcall TxRetDw(double x)
{
    wchar_t tmp[64]; int n = _snwprintf_s(tmp, 64, _TRUNCATE, L"dw:%.0f", x);
    t_wideCounted[0] = static_cast<XCHAR>(n);
    for (int i = 0; i < n; i++) t_wideCounted[i + 1] = tmp[i];
    return t_wideCounted;
}

extern "C" __declspec(dllexport) double* __stdcall TxRetE(double x)
{ t_double = x * 5.0; return &t_double; }

extern "C" __declspec(dllexport) FP12* __stdcall TxRetKw(double x)
{
    static __declspec(thread) unsigned char buf[sizeof(FP12) + 8 * sizeof(double)];
    FP12* f = reinterpret_cast<FP12*>(buf);
    f->rows = 2; f->columns = 2;
    f->array[0] = x; f->array[1] = x + 1; f->array[2] = x + 2; f->array[3] = x + 3;
    return f;
}

// P: returns the NARROW XLOPER. The case the shipped decoder read one byte
// past the end of.
extern "C" __declspec(dllexport) LPXLOPER __stdcall TxRetP(double x)
{
    static __declspec(thread) XLOPER r;
    r.xltype = xltypeNum; r.val.num = x * 7.0;
    return &r;
}

// Q: the wide XLOPER12, in each of its payload shapes.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRetQNum(double x)  { return RetNum(x * 11.0); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRetQStr(double x)
{ wchar_t t[64]; _snwprintf_s(t, 64, _TRUNCATE, L"q:%.0f", x); return RetStr(t); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRetQBool(double x) { return RetBool(x != 0); }
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRetQErr(double)    { return RetErr(xlerrNA); }

// An array return: contents are readable here, and the trace should say so.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRetQArray(double x)
{
    static __declspec(thread) XLOPER12 cells[4];
    cells[0].xltype = xltypeNum;  cells[0].val.num = x;
    cells[1].xltype = xltypeStr;  cells[1].val.str = nullptr;
    cells[2].xltype = xltypeBool; cells[2].val.xbool = 1;
    cells[3].xltype = xltypeErr;  cells[3].val.err = xlerrNA;

    static __declspec(thread) XCHAR two[8];
    two[0] = 3; two[1] = L't'; two[2] = L'w'; two[3] = L'o';
    cells[1].val.str = two;

    t_ret.xltype = xltypeMulti;
    t_ret.val.array.rows = 2;
    t_ret.val.array.columns = 2;
    t_ret.val.array.lparray = cells;
    return &t_ret;
}

// ---- a thread of the add-in's own, calling one of its exports ----------------
//
// Excel is not the only caller of an XLL. Disarm has to be safe while a thread
// the tracer does not know about is inside a detour.
extern "C" __declspec(dllexport) double __stdcall TxHammered(double x) { return x + 1.0; }

namespace
{
    volatile LONG   g_hammerStop  = 0;
    volatile LONG64 g_hammerCalls = 0;
    HANDLE          g_hammer      = nullptr;

    DWORD WINAPI HammerProc(LPVOID)
    {
        // Through the export, as an outside caller would, so nothing inlines past the detour.
        using Fn = double(__stdcall*)(double);
        const Fn f = reinterpret_cast<Fn>(GetProcAddress(ThisModule(), "TxHammered"));
        if (f == nullptr) return 1;
        while (InterlockedCompareExchange(&g_hammerStop, 0, 0) == 0)
        {
            f(1.0);
            InterlockedIncrement64(&g_hammerCalls);
        }
        return 0;
    }
}

extern "C" __declspec(dllexport) int __stdcall TxHammerStart()
{
    if (g_hammer == nullptr)
    {
        InterlockedExchange(&g_hammerStop, 0);
        g_hammer = CreateThread(nullptr, 0, HammerProc, nullptr, 0, nullptr);
    }
    return 1;
}

extern "C" __declspec(dllexport) int __stdcall TxHammerStop()
{
    if (g_hammer != nullptr)
    {
        InterlockedExchange(&g_hammerStop, 1);
        WaitForSingleObject(g_hammer, 10000);
        CloseHandle(g_hammer);
        g_hammer = nullptr;
    }
    return 1;
}

extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxHammerCalls()
{
    return RetNum(static_cast<double>(g_hammerCalls));
}

// ---- registrations made after the tracer armed -------------------------------
//
// Nothing below is registered at load; a command registers it later.

extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxLateNoResult(double x) { return RetNum(x * 7.0); }

// Excel12(xlfRegister, 0, ...): no result asked for, as the SDK's own sample
// registers, so no register id comes back to anyone watching.
extern "C" __declspec(dllexport) int __stdcall TxRegisterNoResult()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);
    PascalStr proc; proc.Set(L"TxLateNoResult");
    PascalStr type; type.Set(L"QB");
    PascalStr func; func.Set(L"TxLateNoResult");
    const int rc = Excel12(xlfRegister, nullptr, 4, &xDLL, &proc.oper, &type.oper, &func.oper);
    Log("register TxLateNoResult with no result rc", rc);
    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}

// TxTwoShapes again, under a WIDER type text: five arguments where arming saw two.
extern "C" __declspec(dllexport) int __stdcall TxRegisterReshape()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);
    Register(xDLL, L"TxTwoShapes", L"QBBBBB", L"TxShapeWide");
    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}

// A procedure name longer than the watch's 127-character capture. The .def
// exports TxPrefixImpl under a 127-character name and TxLongImpl under that
// name plus "Tail", so the cut capture of the long one spells the short one.
extern "C" LPXLOPER12 __stdcall TxPrefixImpl(double x) { return RetNum(x + 0.5); }
extern "C" LPXLOPER12 __stdcall TxLongImpl(double x)   { return RetNum(x + 0.75); }

namespace
{
    constexpr int kPrefixLen = 127;
    void PrefixExportName(wchar_t* out, size_t cap)
    {
        wcscpy_s(out, cap, L"TxPrefix_");
        for (int i = 9; i < kPrefixLen; ++i) out[i] = L'x';
        out[kPrefixLen] = 0;
    }
}

extern "C" __declspec(dllexport) int __stdcall TxRegisterLongName()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);
    wchar_t name[160];
    PrefixExportName(name, _countof(name));
    wcscat_s(name, L"Tail");
    PascalStr proc; proc.Set(name);
    PascalStr type; type.Set(L"QB");
    PascalStr func; func.Set(L"TxLongNamed");
    XLOPER12 res; ZeroMemory(&res, sizeof(res));
    const int rc = Excel12(xlfRegister, &res, 4, &xDLL, &proc.oper, &type.oper, &func.oper);
    Log("register TxLongNamed rc", rc);
    Excel12(xlFree, nullptr, 1, &res);
    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}

// Calls the 127-character export directly, so a detour wrongly placed on it shows in the trace.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxCallsPrefix(double x)
{
    wchar_t w[160];
    PrefixExportName(w, _countof(w));
    char name[160];
    for (int i = 0; i <= kPrefixLen; ++i) name[i] = static_cast<char>(w[i]);
    using Fn = LPXLOPER12(__stdcall*)(double);
    const Fn f = reinterpret_cast<Fn>(GetProcAddress(ThisModule(), name));
    if (f == nullptr) return RetErr(xlerrNA);
    const double v = f(x)->val.num;
    return RetNum(v);
}

// ============================================================================

extern "C" __declspec(dllexport) int __stdcall xlAutoOpen()
{
    XLOPER12 xDLL;
    ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);
    Log("xlAutoOpen xlGetName type", xDLL.xltype);

    // --- argument shapes ---
    Register(xDLL, L"TxB",        L"QBB");
    Register(xDLL, L"TxStackArgs",       L"QBBBBB");     // the fifth is a STACK argument
    Register(xDLL, L"TxA",        L"QA");

    // A display name the other XLL cannot know. Every other function registers under its own
    // export name, so recovering one proves nothing. TxHiddenName appears nowhere except in
    // Excel's own memory, so reading it back proves Excel's registration record was found.
    Register(xDLL, L"TxUnregistered", L"QA", L"TxHiddenName");
    Register(xDLL, L"TxI",        L"QI");
    Register(xDLL, L"TxJ",        L"QJ");
    Register(xDLL, L"TxC",        L"QC");
    Register(xDLL, L"TxD",        L"QD");
    Register(xDLL, L"TxCw",       L"QC%");
    Register(xDLL, L"TxDw",       L"QD%");
    Register(xDLL, L"TxE",        L"QE");
    Register(xDLL, L"TxM",        L"QM");
    Register(xDLL, L"TxN",        L"QN");
    Register(xDLL, L"TxK",        L"QK");
    Register(xDLL, L"TxKw",       L"QK%");
    Register(xDLL, L"TxOw",       L"QO%");        // ONE code, THREE ABI slots
    Register(xDLL, L"TxP",        L"QP");
    Register(xDLL, L"TxQ",        L"QQ");
    Register(xDLL, L"TxMixed",    L"QD%K%E");     // reads as 5 args if parsed per character
    Register(xDLL, L"TxOptional", L"QQQ");

    // the same functions, reached through a jump table instead of a direct
    // export -- a second binary shape, not a second kind of add-in
    Register(xDLL, L"TxJumpAdd", L"QBB");
    Register(xDLL, L"TxJumpStr", L"QD%");

    Register(xDLL, L"TxCallsBack",  L"QB");
    Register(xDLL, L"TxCallsBack2", L"QB");
    Register(xDLL, L"TxMany",       L"QBBBBBBBBBBBB");   // 8 stack arguments
    Register(xDLL, L"TxMacroCalls", L"Q");
    Register(xDLL, L"TxNoArgs",     L"Q");
    Register(xDLL, L"TxRaises",     L"QB");
    Register(xDLL, L"TxDivZero",    L"QB");
    Register(xDLL, L"TxSlow",       L"QB$");   // thread-safe: engages MTC

    // A COMMAND, not a function: macro_type 2 is the sixth xlfRegister
    // argument, so this one cannot go through Register() above.
    {
        PascalStr proc; proc.Set(L"TxMacro");
        PascalStr type; type.Set(L"J");
        PascalStr func; func.Set(L"TxMacro");
        PascalStr args; args.Set(L"");
        XLOPER12 macroType; ZeroMemory(&macroType, sizeof(macroType));
        macroType.xltype = xltypeInt; macroType.val.w = 2;
        XLOPER12 res; ZeroMemory(&res, sizeof(res));
        const int rc = Excel12(xlfRegister, &res, 6, &xDLL, &proc.oper, &type.oper,
                               &func.oper, &args.oper, &macroType);
        Log("register TxMacro (J, command) rc", rc);
        Excel12(xlFree, nullptr, 1, &res);
    }

    // --- return shapes ---
    Register(xDLL, L"TxRetB",     L"BB");         // XMM0, not rax
    Register(xDLL, L"TxRetI",     L"IB");
    Register(xDLL, L"TxRetJ",     L"JB");
    Register(xDLL, L"TxRetC",     L"CB");
    Register(xDLL, L"TxRetD",     L"DB");
    Register(xDLL, L"TxRetCw",    L"C%B");
    Register(xDLL, L"TxRetDw",    L"D%B");
    Register(xDLL, L"TxRetE",     L"EB");
    Register(xDLL, L"TxRetKw",    L"K%B");
    Register(xDLL, L"TxRetP",     L"PB");         // the NARROW XLOPER
    Register(xDLL, L"TxRetQNum",  L"QB");
    Register(xDLL, L"TxRetQStr",  L"QB");
    Register(xDLL, L"TxRetQBool", L"QB");
    Register(xDLL, L"TxRetQErr",  L"QB");
    Register(xDLL, L"TxRetQArray",L"QB");

    // A thread-safe one, so multithreaded calculation actually engages.
    Register(xDLL, L"TxThreadSafe", L"QB$");
    Register(xDLL, L"TxRowCol",     L"QBB$");
    Register(xDLL, L"TxRowColQ",    L"QQQ$");
    Register(xDLL, L"TxRefU",       L"QU");
    Register(xDLL, L"TxRefR",       L"QR");

    // The same export under a second name: arming must hook it once.
    Register(xDLL, L"TxB",           L"QBB", L"TxBAgain");

    Register(xDLL, L"TxTwoShapes",   L"QBB",    L"TxShapeNarrow");

    Register(xDLL, L"TxHammered",    L"BB");
    Register(xDLL, L"TxHammerCalls", L"Q");
    Register(xDLL, L"TxCallsPrefix", L"QB");
    RegisterCommand(xDLL, L"TxHammerStart");
    RegisterCommand(xDLL, L"TxHammerStop");
    RegisterCommand(xDLL, L"TxRegisterNoResult");
    RegisterCommand(xDLL, L"TxRegisterLongName");
    RegisterCommand(xDLL, L"TxRegisterReshape");

    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}

// Returns the cell count of the reference, so a test can check the trace.
namespace
{
    double CellsIn(const XLREF12& r) { return double(r.rwLast - r.rwFirst + 1) * double(r.colLast - r.colFirst + 1); }
    double CellsIn(const XLREF& r)   { return double(r.rwLast - r.rwFirst + 1) * double(r.colLast - r.colFirst + 1); }
}

extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRefU(LPXLOPER12 x)
{
    if (x == nullptr) return RetErr(xlerrValue);
    switch (x->xltype & ~(xlbitXLFree | xlbitDLLFree))
    {
    case xltypeSRef:
        return RetNum(CellsIn(x->val.sref.ref));
    case xltypeRef:
    {
        const XLMREF12* m = x->val.mref.lpmref;
        if (m == nullptr) return RetErr(xlerrValue);
        double n = 0;
        for (WORD i = 0; i < m->count; ++i) n += CellsIn(m->reftbl[i]);
        return RetNum(n);
    }
    default:
        return RetErr(xlerrValue);
    }
}

// The same for the small grid XLOPER.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRefR(LPXLOPER x)
{
    if (x == nullptr) return RetErr(xlerrValue);
    switch (x->xltype & ~(xlbitXLFree | xlbitDLLFree))
    {
    case xltypeSRef:
        return RetNum(CellsIn(x->val.sref.ref));
    case xltypeRef:
    {
        const XLMREF* m = x->val.mref.lpmref;
        if (m == nullptr) return RetErr(xlerrValue);
        double n = 0;
        for (WORD i = 0; i < m->count; ++i) n += CellsIn(m->reftbl[i]);
        return RetNum(n);
    }
    default:
        return RetErr(xlerrValue);
    }
}

// Registered thread-safe ($), so Excel may run it on a worker thread.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxThreadSafe(double x)
{ return RetNum(x * 2.0); }

// Thread-safe: "row:col" from =TxRowCol(ROW(),COLUMN()), so the soak can check every call against its cell.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRowCol(double r, double c)
{
    wchar_t b[64];
    swprintf_s(b, 64, L"%.15g:%.15g", r, c);
    return RetStr(b);
}

namespace
{
    // True for a natural number (1, 2, 3, ...), however Excel passed it: a double, a small
    // integer, or a 1x1 array of one. Otherwise `why` says what is wrong, naming the argument.
    bool Natural(const wchar_t* name, LPXLOPER12 x, double& out, wchar_t* why, size_t cap)
    {
        LPXLOPER12 v = x;
        if (v != nullptr && (v->xltype & ~(xlbitXLFree | xlbitDLLFree)) == xltypeMulti)
        {
            if (v->val.array.rows != 1 || v->val.array.columns != 1)
            {
                swprintf_s(why, cap, L"%s is an array of %d x %d, not a number", name,
                           v->val.array.rows, v->val.array.columns);
                return false;
            }
            v = v->val.array.lparray;
        }
        switch (v != nullptr ? (v->xltype & ~(xlbitXLFree | xlbitDLLFree)) : 0)
        {
        case xltypeNum:     out = v->val.num; break;
        case xltypeInt:     out = v->val.w;   break;
        case xltypeStr:     swprintf_s(why, cap, L"%s is text, not a number", name);            return false;
        case xltypeBool:    swprintf_s(why, cap, L"%s is TRUE or FALSE, not a number", name);   return false;
        case xltypeErr:     swprintf_s(why, cap, L"%s is an error value, not a number", name);  return false;
        case xltypeMissing: swprintf_s(why, cap, L"%s is missing", name);                       return false;
        case xltypeNil:     swprintf_s(why, cap, L"%s is empty", name);                         return false;
        default:            swprintf_s(why, cap, L"%s is not a number", name);                  return false;
        }
        if (!(out >= 1.0) || out != std::floor(out) || out > 9007199254740992.0)
        {
            swprintf_s(why, cap, L"%s %.15g is not a natural number", name, out);
            return false;
        }
        return true;
    }
}

// TxRowCol with XLOPER12 arguments and result. The arguments are Excel's and read-only; the
// result is allocated per call and freed by Excel, so no two threads share a buffer.
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall TxRowColQ(LPXLOPER12 r, LPXLOPER12 c)
{
    double rv = 0, cv = 0;
    wchar_t rowWhy[128] = L"", colWhy[128] = L"", b[300];
    const bool rowOk = Natural(L"row", r, rv, rowWhy, _countof(rowWhy));
    const bool colOk = Natural(L"column", c, cv, colWhy, _countof(colWhy));
    if (rowOk && colOk)        swprintf_s(b, _countof(b), L"%.15g:%.15g", rv, cv);
    else if (!rowOk && !colOk) swprintf_s(b, _countof(b), L"%s; %s", rowWhy, colWhy);
    else                       swprintf_s(b, _countof(b), L"%s", rowOk ? colWhy : rowWhy);
    return RetOwnedStr(b);
}

BOOL APIENTRY DllMain(HMODULE, DWORD, LPVOID) { return TRUE; }
