#pragma once

#include <windows.h>

// Text as Office draws it: DirectWrite, greyscale, hinted. Both DLLs load on first use.
namespace ui
{
namespace text
{
    enum class Face { Body, Bold, Title };

    constexpr unsigned kCentre  = 1;   // horizontally
    constexpr unsigned kVCentre = 2;
    constexpr unsigned kPrefix  = 4;   // '&' underlines the next character
    constexpr unsigned kWrap    = 8;
    constexpr unsigned kUnderline = 16;

    // Whether Direct2D and DirectWrite loaded. There is no other way to draw text here.
    bool Ready();

    bool Draw(HDC dc, const RECT& rc, const wchar_t* s, Face face, COLORREF colour,
              unsigned flags, int dpi, float dx = 0.0f);   // dx: a sub-pixel nudge
    bool Measure(const wchar_t* s, Face face, unsigned flags, int dpi, SIZE& out);

    void Release();
}
}
