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
    constexpr unsigned kEllipsis  = 32;   // text too long for its box ends in "..."

    // Whether Direct2D and DirectWrite loaded. There is no other way to draw text here.
    bool Ready();

    bool Draw(HDC dc, const RECT& rc, const wchar_t* s, Face face, COLORREF colour,
              unsigned flags, int dpi, float dx = 0.0f);   // dx: a sub-pixel nudge

    // Many pieces of text sharing one bind and one flush, which Draw does per call. Every GDI
    // stroke belongs before the batch opens: what it draws reaches the DC only on Close.
    class Batch
    {
    public:
        Batch(HDC dc, const RECT& area);
        ~Batch();
        bool Ok() const { return m_ok; }

        // `rc` and `clip` are in the same coordinates as `area`.
        bool Put(const RECT& rc, const RECT& clip, const wchar_t* s, Face face,
                 COLORREF colour, unsigned flags, int dpi);

    private:
        RECT m_area{};
        bool m_ok = false;
    };
    bool Measure(const wchar_t* s, Face face, unsigned flags, int dpi, SIZE& out);

    void Release();
}
}
