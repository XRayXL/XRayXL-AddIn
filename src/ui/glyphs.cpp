#include "glyphs.h"
#include "softdraw.h"

#include <algorithm>

namespace ui
{
namespace glyph
{
void Draw(HDC dc, const RECT& box, Kind kind, int line, COLORREF ink, COLORREF paper)
{
    const int side = (std::min)(box.right - box.left, box.bottom - box.top);
    // everything is laid out on a 28-unit square
    const auto X = [&](int u) { return box.left + MulDiv(u, side, 28); };
    const auto Y = [&](int u) { return box.top  + MulDiv(u, side, 28); };
    const auto stroke = [&](int x0, int y0, int x1, int y1, double w) {
        const POINT p[2] = { { X(x0), Y(y0) }, { X(x1), Y(y1) } };
        soft::Stroke(dc, p, 2, w, ink); };
    const auto ring = [&](int l, int t, int r, int b, int radius) {
        const RECT rc{ X(l), Y(t), X(r), Y(b) };
        soft::RoundRect(dc, rc, soft::Box{ 0, ink, line, MulDiv(radius, side, 28), false }); };
    const auto disc = [&](int cx, int cy, int radius, COLORREF c) {
        const RECT rc{ X(cx - radius), Y(cy - radius), X(cx + radius), Y(cy + radius) };
        soft::RoundRect(dc, rc, soft::Box{ c, c, 0, MulDiv(radius, side, 28), true }); };

    switch (kind)
    {
    case Kind::Capture:                     // a list with a record dot
        ring(1, 3, 23, 21, 2);
        stroke(5, 8, 7, 8, line);  stroke(10, 8, 18, 8, line);
        stroke(5, 12, 7, 12, line); stroke(10, 12, 15, 12, line);
        stroke(5, 16, 7, 16, line);
        disc(21, 20, 7, paper);
        ring(15, 14, 27, 26, 6);
        disc(21, 20, 3, ink);
        break;
    case Kind::Events:                      // a lightning bolt
        stroke(17, 2, 8, 15, line);  stroke(8, 15, 14, 15, line);
        stroke(14, 15, 11, 26, line); stroke(11, 26, 21, 12, line);
        stroke(21, 12, 15, 12, line); stroke(15, 12, 17, 2, line);
        break;
    case Kind::Output:                      // a page, with an arrow leaving it
        ring(3, 1, 19, 25, 2);
        stroke(7, 7, 15, 7, line);
        stroke(7, 11, 15, 11, line);
        stroke(7, 15, 12, 15, line);
        disc(21, 20, 7, paper);
        stroke(14, 20, 26, 20, line * 1.2);
        stroke(22, 16, 26, 20, line * 1.2);
        stroke(22, 24, 26, 20, line * 1.2);
        break;
    case Kind::Advanced:                    // three sliders
        for (int i = 0; i < 3; ++i)
        {
            const int y = 6 + i * 8, knob = (i == 0) ? 9 : (i == 1) ? 19 : 13;
            stroke(1, y, 27, y, line);
            disc(knob, y, 4, paper);
            ring(knob - 3, y - 3, knob + 4, y + 4, 4);      // odd width: centred on the line's own pixel row
        }
        break;
    case Kind::About:                       // an i in a circle, all on the axis x = 14
    {
        ring(1, 1, 27, 27, 13);
        const double u = side / 28.0;
        soft::Disc(dc, box.left + 14 * u, box.top + 8.5 * u, 1.7 * u, ink);
        const RECT stem{ X(13), Y(12), X(15), Y(21) };
        soft::RoundRect(dc, stem, soft::Box{ ink, ink, 0, 1, true });
        break;
    }
    case Kind::Notices:                     // a page of text
        ring(4, 1, 24, 27, 2);
        stroke(8, 7, 20, 7, line);
        stroke(8, 11, 20, 11, line);
        stroke(8, 15, 20, 15, line);
        stroke(8, 19, 16, 19, line);
        break;
    case Kind::Modules:                     // three stacked blocks, as a loaded module list
        for (int i = 0; i < 3; ++i)
        {
            const int y = 2 + i * 9;
            ring(2, y, 26, y + 7, 2);
            stroke(5, y + 3, 9, y + 3, line);
            stroke(12, y + 3, 22, y + 3, line);
        }
        break;
    case Kind::Environment:                 // a name and its value, twice, with an = between
        stroke(2, 8, 9, 8, line);
        stroke(12, 6, 18, 6, line);  stroke(12, 10, 18, 10, line);
        stroke(21, 8, 26, 8, line);
        stroke(2, 20, 9, 20, line);
        stroke(12, 18, 18, 18, line); stroke(12, 22, 18, 22, line);
        stroke(21, 20, 26, 20, line);
        break;
    case Kind::Process:                     // a chip: a square with legs on every side
        ring(7, 7, 21, 21, 2);
        ring(11, 11, 17, 17, 1);
        for (int i = 0; i < 3; ++i)
        {
            const int at = 10 + i * 4;
            stroke(at, 2, at, 7, line);
            stroke(at, 21, at, 26, line);
            stroke(2, at, 7, at, line);
            stroke(21, at, 26, at, line);
        }
        break;
    }
}
}
}
