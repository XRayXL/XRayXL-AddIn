#pragma once
#include <windows.h>
#include <cstdint>

// QueryPerformanceCounter as microseconds, for the arm-time cost splits. The hot path stamps raw
// ticks and converts nothing.
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

    // The time this thread has spent inside the tracers' own hooks, shared by both sources so a
    // VBA call's `tracerticks` includes the XLL calls it made. A frame keeps its value at entry.
    inline std::uint64_t& TracerTicks()
    {
        static thread_local std::uint64_t ticks = 0;
        return ticks;
    }

    inline long long QpcMicros()
    {
        LARGE_INTEGER f, t;
        QueryPerformanceFrequency(&f);
        QueryPerformanceCounter(&t);
        return QpcMicrosFrom(t.QuadPart, f.QuadPart);
    }
}
