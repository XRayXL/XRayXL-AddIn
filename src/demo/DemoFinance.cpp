// DemoFinance64.xll -- numeric / financial worksheet functions for playing with
// XRayXL. Between them they show the timing column doing real work: a fast
// closed-form price, a deliberately slow one, and a compute-heavy recursion.
#include "demo_xll_common.h"
#include <cmath>

using namespace demo;

namespace
{
    double NormCdf(double x) { return 0.5 * (1.0 + std::erf(x / std::sqrt(2.0))); }

    // Naive recursive Fibonacci -- exponential on purpose, so Fib(30+) is a
    // procedure that visibly TAKES time in the trace, not an instant return.
    double FibRec(double n)
    {
        if (n < 2.0) return n;
        return FibRec(n - 1.0) + FibRec(n - 2.0);
    }
}

// A European call price (Black-Scholes). Fast, pure math -- the "instant" end of
// the timing column. =BlackScholes(100, 100, 1, 0.05, 0.2)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
BlackScholes(double spot, double strike, double years, double rate, double vol)
{
    if (spot <= 0 || strike <= 0 || years <= 0 || vol <= 0) return RetErr(xlerrValue);
    const double sqrtT = std::sqrt(years);
    const double d1 = (std::log(spot / strike) + (rate + 0.5 * vol * vol) * years) / (vol * sqrtT);
    const double d2 = d1 - vol * sqrtT;
    return RetNum(spot * NormCdf(d1) - strike * std::exp(-rate * years) * NormCdf(d2));
}

// Present value of a single cashflow. =PresentValue(1000, 0.05, 10)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
PresentValue(double cashflow, double rate, double periods)
{
    return RetNum(cashflow / std::pow(1.0 + rate, periods));
}

// Compound growth. =CompoundReturn(1000, 0.07, 30)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
CompoundReturn(double principal, double rate, double years)
{
    return RetNum(principal * std::pow(1.0 + rate, years));
}

// Deliberately SLOW: ~40ms of nothing, then the sum. The obvious culprit when
// you sort a trace by duration. =SlowSum(2, 3)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
SlowSum(double a, double b)
{
    Sleep(40);
    return RetNum(a + b);
}

// Compute-heavy: naive recursion, capped so it cannot hang Excel. Fib(30) is a
// noticeable pause; the trace shows one long XLL call. =Fibonacci(30)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
Fibonacci(double n)
{
    if (n < 0) return RetErr(xlerrValue);
    if (n > 38) n = 38;          // ~40M calls; keep the demo responsive
    return RetNum(FibRec(std::floor(n)));
}

namespace
{
    // The numbers in a Q argument: one cell, or a range's cells row by row. Other cells are skipped.
    int NumbersIn(LPXLOPER12 x, double* out, int cap)
    {
        const DWORD t = x->xltype & ~(xlbitXLFree | xlbitDLLFree);
        if (t == xltypeNum) { if (cap > 0) out[0] = x->val.num; return 1; }
        if (t != xltypeMulti) return 0;
        const int cells = x->val.array.rows * x->val.array.columns;
        int n = 0;
        for (int i = 0; i < cells && n < cap; ++i)
            if (x->val.array.lparray[i].xltype == xltypeNum) out[n++] = x->val.array.lparray[i].val.num;
        return n;
    }

    // A counted Excel string against a null-terminated one, ignoring case.
    bool SameText(const XCHAR* counted, const XCHAR* z)
    {
        const size_t len = counted[0];
        if (wcslen(z) != len) return false;
        return _wcsnicmp(counted + 1, z, len) == 0;
    }
}

// Net present value of a row or column of yearly cashflows, the first a year out.
// A range passed as Q arrives as its values. =NetPresentValue(0.05, B2:F2)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
NetPresentValue(double rate, LPXLOPER12 cashflows)
{
    double cf[256];
    const int n = NumbersIn(cashflows, cf, 256);
    if (n == 0) return RetErr(xlerrValue);
    double pv = 0;
    for (int i = 0; i < n; ++i) pv += cf[i] / std::pow(1.0 + rate, i + 1);
    return RetNum(pv);
}

// A discount factor for each tenor in years, in the tenors' own shape. =DiscountCurve(0.05, B2:B6)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
DiscountCurve(double rate, LPXLOPER12 years)
{
    const DWORD t = years->xltype & ~(xlbitXLFree | xlbitDLLFree);
    if (t == xltypeNum) return RetNum(std::pow(1.0 + rate, -years->val.num));
    if (t != xltypeMulti) return RetErr(xlerrValue);
    const int rows = years->val.array.rows, cols = years->val.array.columns;
    if (rows * cols > 256) return RetErr(xlerrNum);
    for (int i = 0; i < rows * cols; ++i)
    {
        const XLOPER12& y = years->val.array.lparray[i];
        if (y.xltype == xltypeNum) { t_arr[i].xltype = xltypeNum; t_arr[i].val.num = std::pow(1.0 + rate, -y.val.num); }
        else                       { t_arr[i].xltype = xltypeErr; t_arr[i].val.err = xlerrValue; }
    }
    t_ret.xltype = xltypeMulti;
    t_ret.val.array.rows = rows;
    t_ret.val.array.columns = cols;
    t_ret.val.array.lparray = t_arr;
    return &t_ret;
}

// Average of values weighted by weights, both plain arrays of doubles (K%, FP12).
// =WeightedAverage(C2:C11, D2:D11)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
WeightedAverage(FP12* values, FP12* weights)
{
    const int n = values->rows * values->columns;
    if (n != weights->rows * weights->columns) return RetErr(xlerrValue);
    double sum = 0, wsum = 0;
    for (int i = 0; i < n; ++i) { sum += values->array[i] * weights->array[i]; wsum += weights->array[i]; }
    if (wsum == 0) return RetErr(xlerrDiv0);
    return RetNum(sum / wsum);
}

// How many cells a reference covers, read from the reference itself (U), never its values.
// =RangeSize(A2:E11)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
RangeSize(LPXLOPER12 ref)
{
    const DWORD t = ref->xltype & ~(xlbitXLFree | xlbitDLLFree);
    double cells = 0;
    if (t == xltypeSRef)
    {
        const XLREF12& r = ref->val.sref.ref;
        cells = double(r.rwLast - r.rwFirst + 1) * double(r.colLast - r.colFirst + 1);
    }
    else if (t == xltypeRef)
    {
        const XLMREF12* m = ref->val.mref.lpmref;
        for (WORD i = 0; m && i < m->count; ++i)
            cells += double(m->reftbl[i].rwLast - m->reftbl[i].rwFirst + 1) *
                     double(m->reftbl[i].colLast - m->reftbl[i].colFirst + 1);
    }
    else return RetNum(1);   // a value, not a reference
    return RetNum(cells);
}

// The price beside a ticker in a two-column table, or #N/A. =QuoteLookup("ACME", A2:B6)
extern "C" __declspec(dllexport) LPXLOPER12 __stdcall
QuoteLookup(const XCHAR* ticker, LPXLOPER12 table)
{
    const DWORD t = table->xltype & ~(xlbitXLFree | xlbitDLLFree);
    if (t != xltypeMulti || table->val.array.columns < 2) return RetErr(xlerrValue);
    const int cols = table->val.array.columns;
    for (int r = 0; r < table->val.array.rows; ++r)
    {
        const XLOPER12& key = table->val.array.lparray[r * cols];
        const XLOPER12& val = table->val.array.lparray[r * cols + 1];
        if (key.xltype == xltypeStr && SameText(key.val.str, ticker))
            return val.xltype == xltypeNum ? RetNum(val.val.num) : RetErr(xlerrValue);
    }
    return RetErr(xlerrNA);
}

extern "C" __declspec(dllexport) int __stdcall xlAutoOpen()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);

    Register(xDLL, L"BlackScholes",   L"QBBBBB");
    Register(xDLL, L"PresentValue",   L"QBBB");
    Register(xDLL, L"CompoundReturn", L"QBBB");
    Register(xDLL, L"SlowSum",        L"QBB");
    Register(xDLL, L"Fibonacci",      L"QB");
    Register(xDLL, L"NetPresentValue", L"QBQ");     // a range, as its values
    Register(xDLL, L"DiscountCurve",   L"QBQ");     // an array result in the input's shape
    Register(xDLL, L"WeightedAverage", L"QK%K%");   // plain double arrays
    Register(xDLL, L"RangeSize",       L"QU");      // a reference, not its values
    Register(xDLL, L"QuoteLookup",     L"QC%Q");    // text and a table

    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}
