#pragma once
#include <windows.h>

// QueryPerformanceCounter as microseconds, for the arm-time cost splits that
// say where a slow arm went. Never GetTickCount. Not for the hot path, which
// stamps raw ticks and converts nothing.
namespace core
{
    // Raw ticks, for the hot path.
    inline long long QpcNow()
    {
        LARGE_INTEGER t;
        QueryPerformanceCounter(&t);
        return t.QuadPart;
    }

    // Whole seconds first: multiplying the raw counter by a million overflows after
    // about ten days at a 10 MHz counter, and within an hour at a TSC rate.
    inline long long QpcMicrosFrom(long long ticks, long long frequency)
    {
        return (ticks / frequency) * 1000000LL + ((ticks % frequency) * 1000000LL) / frequency;
    }

    inline long long QpcMicros()
    {
        LARGE_INTEGER f, t;
        QueryPerformanceFrequency(&f);
        QueryPerformanceCounter(&t);
        return QpcMicrosFrom(t.QuadPart, f.QuadPart);
    }
}
