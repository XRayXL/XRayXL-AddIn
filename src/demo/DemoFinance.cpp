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

extern "C" __declspec(dllexport) int __stdcall xlAutoOpen()
{
    XLOPER12 xDLL; ZeroMemory(&xDLL, sizeof(xDLL));
    Excel12(xlGetName, &xDLL, 0);

    Register(xDLL, L"BlackScholes",   L"QBBBBB");
    Register(xDLL, L"PresentValue",   L"QBBB");
    Register(xDLL, L"CompoundReturn", L"QBBB");
    Register(xDLL, L"SlowSum",        L"QBB");
    Register(xDLL, L"Fibonacci",      L"QB");

    Excel12(xlFree, nullptr, 1, &xDLL);
    return 1;
}
