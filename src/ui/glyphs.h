#pragma once

#include <windows.h>

// The dialog's page glyphs: hairline drawings in one ink, soft at any size.

namespace ui
{
namespace glyph
{
    enum class Kind { Capture, Output, Advanced, About, Notices };

    // Draws `kind` in the square at the top-left of `box`; `paper` is the colour behind it.
    void Draw(HDC dc, const RECT& box, Kind kind, int linePx, COLORREF ink, COLORREF paper);
}
}
