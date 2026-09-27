// The ribbon's button pictures: a page, most with a round badge over its lower right holding the
// button's mark. Drawn on a 32-unit grid, a large button's icon at 96 DPI, and scaled to the screen's DPI.

#include "ribbonart.h"

#include <olectl.h>
#include <cmath>
#include <vector>

namespace ui
{
namespace ribbon
{
namespace art
{
namespace
{
    // As a 32-bit DIB holds a pixel: blue first, colour premultiplied by alpha.
    struct Pixel { BYTE b, g, r, a; };

    const Pixel kInk   { 0x38, 0x3A, 0x3A, 0xFF };
    const Pixel kGrey  { 0x74, 0x77, 0x79, 0xFF };      // what a page lists
    const Pixel kSheet { 0xFA, 0xFA, 0xFA, 0xFF };      // inside the badge, and the checklist's paper
    const Pixel kGreen { 0x41, 0x7C, 0x10, 0xFF };      // Excel's green, #107C41
    const Pixel kRed   { 0x4F, 0x4C, 0xE9, 0xFF };      // Office's record red, #E94C4F
    const Pixel kClear { 0x00, 0x00, 0x00, 0x00 };

    // ---- shapes ------------------------------------------------------------------------------

    // Includes its left and top edges, not its right and bottom.
    struct Box { double left, top, right, bottom; };

    bool In(const Box& b, double x, double y)
    {
        return x >= b.left && x < b.right && y >= b.top && y < b.bottom;
    }

    template <size_t N>
    bool In(const Box (&boxes)[N], double x, double y)
    {
        for (const Box& b : boxes)
            if (In(b, x, y)) return true;
        return false;
    }

    // The page behind Arm, Disarm, Tail and Perfetto: a window with a title bar, holding a list.
    const Box kWindowFrame[] = {
        {  2,  3, 30,  4 }, {  2, 28, 30, 29 },             // top, bottom
        {  2,  3,  3, 29 }, { 29,  3, 30, 29 },             // left, right
        {  2,  7, 30,  8 },                                 // under the title bar
    };
    const Box kWindowList[] = {
        {  6, 11,  8, 12 }, { 11, 11, 16, 12 }, { 18, 11, 26, 12 },
        {  6, 15,  8, 16 }, { 11, 15, 17, 16 }, { 19, 15, 22, 16 },
        {  6, 19,  8, 20 }, { 11, 19, 17, 20 }, { 19, 19, 27, 20 },
        {  6, 23,  8, 24 }, { 11, 23, 15, 24 }, { 18, 23, 24, 24 },
    };

    // The page behind Diagnostics: a sheet of paper with its top right corner folded down.
    const Box kPaperEdges[] = {
        {  5,  2,  6, 30 }, {  5, 29, 26, 30 },             // left, bottom
        {  5,  2, 18,  3 }, { 25, 10, 26, 30 },             // top, right
        { 17,  2, 18, 11 }, { 17, 10, 26, 11 },             // the fold's two straight sides
    };
    // The fold's diagonal, from the top edge's end to the right edge's top, drawn a little heavier
    // than a unit so it weighs the same as the straight edges.
    bool OnFold(double x, double y)
    {
        return x >= 17 && x < 26 && y >= 2 && y < 11 && std::fabs(x - y - 15.25) <= 0.75;
    }

    // The page behind Options: a sheet listing three items, each with a check box, as Office's own.
    const Box kChecklistPaper = { 5, 2, 26, 30 };
    const Box kChecklistEdges[] = {
        {  5,  2, 26,  3 }, {  5, 29, 26, 30 },             // top, bottom
        {  5,  2,  6, 30 }, { 25,  2, 26, 30 },             // left, right
    };
    const Box kChecklistItems[] = {                         // a box's four sides, then its line
        {  9,  7, 13,  8 }, {  9, 10, 13, 11 }, {  9,  7, 10, 11 }, { 12,  7, 13, 11 }, { 15,  8, 22,  9 },
        {  9, 14, 13, 15 }, {  9, 17, 13, 18 }, {  9, 14, 10, 18 }, { 12, 14, 13, 18 }, { 15, 15, 22, 16 },
        {  9, 21, 13, 22 }, {  9, 24, 13, 25 }, {  9, 21, 10, 25 }, { 12, 21, 13, 25 }, { 15, 22, 22, 23 },
    };

    enum class Page { Window, Paper, Checklist };

    const Pixel& PageAt(Page page, double x, double y)
    {
        switch (page)
        {
        case Page::Window:
            if (In(kWindowFrame, x, y)) return kInk;
            return In(kWindowList, x, y) ? kGrey : kClear;
        case Page::Paper:
            return In(kPaperEdges, x, y) || OnFold(x, y) ? kInk : kClear;
        case Page::Checklist:
            if (In(kChecklistEdges, x, y)) return kInk;
            if (In(kChecklistItems, x, y)) return kGrey;
            return In(kChecklistPaper, x, y) ? kSheet : kClear;
        }
        return kClear;
    }

    // ---- the badge and its marks ---------------------------------------------------------------

    // A ring of ink round a sheet holding the mark, with a clear gap outside it cutting the page.
    const double kBadgeX = 21.2, kBadgeY = 21.2, kBadgeRadius = 10.8;
    const double kRingWidth = 32.0 / 30, kGapWidth = 1;

    // Marks are measured in the stop square's side, from the badge's centre.
    const double kMarkSize = 10.08;

    // Three bars: a call, and two nested beneath it, as Perfetto draws them.
    const Box kTimeline[] = {
        { -0.6,  -0.5,  0.6, -0.22 },
        { -0.45, -0.14, 0.3,  0.14 },
        { -0.3,   0.22, 0.0,  0.5  },
    };
    // Two bars, as on Office's Document Inspector.
    const Box kEquals[] = {
        { -0.643, -0.357, 0.643, -0.214 },
        { -0.643,  0.214, 0.643,  0.357 },
    };

    // An equilateral triangle pointing down, its corners as far out as the square's. Centred on its
    // centroid it looks low beside the square, so it sits a tenth higher.
    bool InTriangle(double x, double y)
    {
        const double base = 1.2247, height = base * 0.8660;
        const double top = -height / 3 - 0.1, apex = 2 * height / 3 - 0.1;
        return y >= top && y <= apex && std::fabs(x) <= (base / 2) * (apex - y) / height;
    }

    enum class Mark { None, Record, Stop, Follow, Timeline, Equals };

    // The mark's colour at (x, y) in mark units, or null where the sheet shows.
    const Pixel* MarkAt(Mark mark, double x, double y)
    {
        switch (mark)
        {
        case Mark::None:     return nullptr;
        case Mark::Record:   return x * x + y * y <= 0.25                     ? &kRed   : nullptr;
        case Mark::Stop:     return std::fabs(x) <= 0.5 && std::fabs(y) <= 0.5 ? &kInk   : nullptr;
        case Mark::Follow:   return InTriangle(x, y)                          ? &kGreen : nullptr;
        case Mark::Timeline: return In(kTimeline, x, y)                       ? &kGreen : nullptr;
        case Mark::Equals:   return In(kEquals, x, y)                         ? &kRed   : nullptr;
        }
        return nullptr;
    }

    // ---- drawing -------------------------------------------------------------------------------

    struct Icon { Page page; Mark mark; };

    const Pixel& ColourAt(const Icon& icon, double x, double y)
    {
        const double d = std::hypot(x - kBadgeX, y - kBadgeY);
        if (icon.mark == Mark::None || d > kBadgeRadius + kGapWidth) return PageAt(icon.page, x, y);
        if (d > kBadgeRadius)              return kClear;
        if (d > kBadgeRadius - kRingWidth) return kInk;
        const Pixel* m = MarkAt(icon.mark, (x - kBadgeX) / kMarkSize, (y - kBadgeY) / kMarkSize);
        return m ? *m : kSheet;
    }

    // The icon at `size` pixels square. Each pixel averages 4x4 samples; as every colour is opaque
    // or clear, the average is already premultiplied.
    std::vector<Pixel> Draw(const Icon& icon, int size)
    {
        const double unitsPerPixel = 32.0 / size;
        std::vector<Pixel> px;
        px.reserve(static_cast<size_t>(size) * size);
        for (int y = 0; y < size; ++y)
            for (int x = 0; x < size; ++x)
            {
                int b = 0, g = 0, r = 0, a = 0;
                for (int sy = 0; sy < 4; ++sy)
                    for (int sx = 0; sx < 4; ++sx)
                    {
                        const Pixel& c = ColourAt(icon, (x + (sx + 0.5) / 4) * unitsPerPixel,
                                                        (y + (sy + 0.5) / 4) * unitsPerPixel);
                        b += c.b; g += c.g; r += c.r; a += c.a;
                    }
                px.push_back(Pixel{ static_cast<BYTE>(b / 16), static_cast<BYTE>(g / 16),
                                    static_cast<BYTE>(r / 16), static_cast<BYTE>(a / 16) });
            }
        return px;
    }

    // ---- handing it to Office ------------------------------------------------------------------

    int DpiOf(HWND window)
    {
        // GetDpiForWindow, where Windows has it
        typedef UINT (WINAPI* Fn)(HWND);
        static const auto fn = reinterpret_cast<Fn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow")));
        const UINT dpi = (fn && window) ? fn(window) : 0;
        return dpi ? static_cast<int>(dpi) : 96;
    }

    // An IPictureDisp that owns a bitmap of these pixels.
    IDispatch* PictureOf(const std::vector<Pixel>& px, int size)
    {
        BITMAPINFO bi{};
        bi.bmiHeader.biSize = sizeof(bi.bmiHeader);
        bi.bmiHeader.biWidth = size;
        bi.bmiHeader.biHeight = -size;                  // top-down
        bi.bmiHeader.biPlanes = 1;
        bi.bmiHeader.biBitCount = 32;
        void* bits = nullptr;
        HDC dc = GetDC(nullptr);
        HBITMAP bitmap = CreateDIBSection(dc, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
        ReleaseDC(nullptr, dc);
        if (!bitmap || !bits) { if (bitmap) DeleteObject(bitmap); return nullptr; }
        memcpy(bits, px.data(), px.size() * sizeof(Pixel));

        PICTDESC desc{};
        desc.cbSizeofstruct = sizeof(desc);
        desc.picType = PICTYPE_BITMAP;
        desc.bmp.hbitmap = bitmap;
        IDispatch* picture = nullptr;
        if (FAILED(OleCreatePictureIndirect(&desc, IID_IDispatch, TRUE, reinterpret_cast<void**>(&picture))))
        {
            DeleteObject(bitmap);
            return nullptr;
        }
        return picture;
    }

    IDispatch* Picture(const Icon& icon, HWND dpiOf)
    {
        const int size = MulDiv(32, DpiOf(dpiOf), 96);
        return PictureOf(Draw(icon, size), size);
    }
}

IDispatch* ArmPicture(HWND dpiOf)         { return Picture({ Page::Window,    Mark::Record },   dpiOf); }
IDispatch* DisarmPicture(HWND dpiOf)      { return Picture({ Page::Window,    Mark::Stop },     dpiOf); }
IDispatch* TailPicture(HWND dpiOf)        { return Picture({ Page::Window,    Mark::Follow },   dpiOf); }
IDispatch* PerfettoPicture(HWND dpiOf)    { return Picture({ Page::Window,    Mark::Timeline }, dpiOf); }
IDispatch* OptionsPicture(HWND dpiOf)     { return Picture({ Page::Checklist, Mark::None },     dpiOf); }
IDispatch* DiagnosticsPicture(HWND dpiOf) { return Picture({ Page::Paper,     Mark::Equals },   dpiOf); }
}
}
}
