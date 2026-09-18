#pragma once
#include <windows.h>

// Anti-aliased rounded shapes and strokes, blended over what is on the device.
namespace ui
{
namespace soft
{
    struct Box
    {
        COLORREF fill;        // inside the edge (ignored unless fillInside)
        COLORREF edge;        // the edge band
        int      edgePx;      // 0 for a solid shape in `edge`
        int      radiusPx;    // corner radius; clamped to half the shorter side
        bool     fillInside;  // false = a ring only; what is under it shows through
    };

    // Paints `box` over `rc` on `dc`, blending against what is there.
    void RoundRect(HDC dc, const RECT& rc, const Box& box);

    // An anti-aliased polyline `widthPx` wide: the chevron and the tick.
    void Stroke(HDC dc, const POINT* pts, int count, double widthPx, COLORREF colour);

    // A filled circle; centre and radius need not be whole pixels.
    void Disc(HDC dc, double cx, double cy, double radius, COLORREF colour);

    // Blends `colour` through a w*h map of coverage, 0..255.
    void Plot(HDC dc, int x, int y, int w, int h, const unsigned char* coverage, COLORREF colour);
}
}
