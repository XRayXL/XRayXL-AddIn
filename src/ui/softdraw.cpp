#include "softdraw.h"

#include <algorithm>
#include <cmath>

namespace ui
{
namespace soft
{
namespace
{
    inline double Clamp01(double v) { return v < 0.0 ? 0.0 : (v > 1.0 ? 1.0 : v); }

    // Signed distance to a rounded rectangle; negative inside.
    double RoundedRectDistance(double px, double py, double cx, double cy,
                               double hw, double hh, double r)
    {
        const double qx = std::fabs(px - cx) - (hw - r);
        const double qy = std::fabs(py - cy) - (hh - r);
        const double ax = qx > 0.0 ? qx : 0.0;
        const double ay = qy > 0.0 ? qy : 0.0;
        const double outside = std::sqrt(ax * ax + ay * ay);
        const double inside  = (std::min)((std::max)(qx, qy), 0.0);
        return outside + inside - r;
    }

    // Distance from (px,py) to the segment a-b.
    double SegmentDistance(double px, double py, double ax, double ay, double bx, double by)
    {
        const double vx = bx - ax, vy = by - ay;
        const double wx = px - ax, wy = py - ay;
        const double len2 = vx * vx + vy * vy;
        const double t = len2 > 0.0 ? Clamp01((wx * vx + wy * vy) / len2) : 0.0;
        const double dx = wx - t * vx, dy = wy - t * vy;
        return std::sqrt(dx * dx + dy * dy);
    }

    // One pixel of falloff either side of the edge.
    inline double Coverage(double d) { return Clamp01(0.5 - d); }

    inline BYTE Mix(BYTE from, BYTE to, double t)
    {
        return static_cast<BYTE>(from + (static_cast<int>(to) - from) * t + 0.5);
    }

    // A scratch copy of the device rectangle, so shapes blend against real pixels.
    struct Scratch
    {
        HDC     mem  = nullptr;
        HBITMAP bmp  = nullptr;
        HGDIOBJ old  = nullptr;
        BYTE*   px   = nullptr;
        int     w = 0, h = 0;

        bool Open(HDC dc, const RECT& rc)
        {
            w = rc.right - rc.left; h = rc.bottom - rc.top;
            if (w <= 0 || h <= 0) return false;
            BITMAPINFO bi{};
            bi.bmiHeader.biSize        = sizeof(bi.bmiHeader);
            bi.bmiHeader.biWidth       = w;
            bi.bmiHeader.biHeight      = -h;         // top-down
            bi.bmiHeader.biPlanes      = 1;
            bi.bmiHeader.biBitCount    = 32;
            bi.bmiHeader.biCompression = BI_RGB;
            void* bits = nullptr;
            mem = CreateCompatibleDC(dc);
            bmp = CreateDIBSection(dc, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
            if (!mem || !bmp || !bits) { Close(); return false; }
            old = SelectObject(mem, bmp);
            px  = static_cast<BYTE*>(bits);
            BitBlt(mem, 0, 0, w, h, dc, rc.left, rc.top, SRCCOPY);
            GdiFlush();
            return true;
        }
        void Commit(HDC dc, const RECT& rc)
        {
            GdiFlush();
            BitBlt(dc, rc.left, rc.top, w, h, mem, 0, 0, SRCCOPY);
        }
        void Close()
        {
            if (mem && old) SelectObject(mem, old);
            if (bmp) DeleteObject(bmp);
            if (mem) DeleteDC(mem);
            mem = nullptr; bmp = nullptr; old = nullptr; px = nullptr;
        }
        ~Scratch() { Close(); }

        void Blend(int x, int y, COLORREF c, double t)
        {
            if (t <= 0.0) return;
            BYTE* p = px + (static_cast<size_t>(y) * w + x) * 4;
            p[0] = Mix(p[0], GetBValue(c), t);
            p[1] = Mix(p[1], GetGValue(c), t);
            p[2] = Mix(p[2], GetRValue(c), t);
            p[3] = 0xFF;
        }
    };
}

void RoundRect(HDC dc, const RECT& rc, const Box& box)
{
    Scratch s;
    if (!s.Open(dc, rc)) return;

    const double hw = s.w / 2.0, hh = s.h / 2.0;
    const double cx = hw, cy = hh;
    const double r  = (std::min<double>)(box.radiusPx, (std::min)(hw, hh));
    const double e  = box.edgePx > 0 ? box.edgePx : 0.0;
    const double ri = (std::max)(r - e, 0.0);

    // A heavy edge is two thin rings, as Office draws it.
    const bool   heavy = e >= 2.0 && r > e;
    const double half  = e / 2.0;
    const double rMid  = (std::max)(r - half, 0.0);           // inside of the outer ring
    const double rB    = (std::max)(r - half / 2.0, 0.0);     // outside of the inner ring
    const double rBin  = (std::max)(rB - half, 0.0);          // inside of the inner ring

    for (int y = 0; y < s.h; ++y)
        for (int x = 0; x < s.w; ++x)
        {
            const double sx = x + 0.5, sy = y + 0.5;
            const double cOut = Coverage(RoundedRectDistance(sx, sy, cx, cy, hw, hh, r));
            if (heavy)
            {
                const double cMid = Coverage(RoundedRectDistance(sx, sy, cx, cy, hw - half, hh - half, rMid));
                const double cB   = Coverage(RoundedRectDistance(sx, sy, cx, cy, hw - half, hh - half, rB));
                const double cBin = Coverage(RoundedRectDistance(sx, sy, cx, cy, hw - e, hh - e, rBin));
                s.Blend(x, y, box.edge, box.fillInside ? cOut : (std::max)(cOut - cMid, 0.0));
                if (box.fillInside) s.Blend(x, y, box.fill, cMid);
                s.Blend(x, y, box.edge, (std::max)(cB - cBin, 0.0));
                continue;
            }
            const double cIn  = e > 0.0
                ? Coverage(RoundedRectDistance(sx, sy, cx, cy, hw - e, hh - e, ri))
                : cOut;
            // the edge band is what the outer shape covers and the inner does not
            const double edgeCov = box.fillInside ? cOut : (std::max)(cOut - cIn, 0.0);
            s.Blend(x, y, box.edge, edgeCov);
            if (box.fillInside) s.Blend(x, y, box.fill, cIn);
        }
    s.Commit(dc, rc);
}

void Stroke(HDC dc, const POINT* pts, int count, double widthPx, COLORREF colour)
{
    if (!pts || count < 2) return;
    RECT rc{ pts[0].x, pts[0].y, pts[0].x, pts[0].y };
    for (int i = 1; i < count; ++i)
    {
        rc.left = (std::min)(rc.left, pts[i].x); rc.right  = (std::max)(rc.right,  pts[i].x);
        rc.top  = (std::min)(rc.top,  pts[i].y); rc.bottom = (std::max)(rc.bottom, pts[i].y);
    }
    const int pad = static_cast<int>(std::ceil(widthPx)) + 2;
    rc.left -= pad; rc.top -= pad; rc.right += pad + 1; rc.bottom += pad + 1;

    Scratch s;
    if (!s.Open(dc, rc)) return;

    const double half = widthPx / 2.0;
    for (int y = 0; y < s.h; ++y)
        for (int x = 0; x < s.w; ++x)
        {
            const double sx = rc.left + x + 0.5, sy = rc.top + y + 0.5;
            double d = 1e9;
            for (int i = 0; i + 1 < count; ++i)
                d = (std::min)(d, SegmentDistance(sx, sy,
                                                pts[i].x + 0.5, pts[i].y + 0.5,
                                                pts[i + 1].x + 0.5, pts[i + 1].y + 0.5));
            s.Blend(x, y, colour, Coverage(d - half));
        }
    s.Commit(dc, rc);
}

void Disc(HDC dc, double cx, double cy, double radius, COLORREF colour)
{
    const RECT rc{ static_cast<LONG>(std::floor(cx - radius)) - 1, static_cast<LONG>(std::floor(cy - radius)) - 1,
                   static_cast<LONG>(std::ceil(cx + radius)) + 1,  static_cast<LONG>(std::ceil(cy + radius)) + 1 };
    Scratch s;
    if (!s.Open(dc, rc)) return;
    for (int y = 0; y < s.h; ++y)
        for (int x = 0; x < s.w; ++x)
        {
            const double dx = rc.left + x + 0.5 - cx, dy = rc.top + y + 0.5 - cy;
            s.Blend(x, y, colour, Coverage(std::sqrt(dx * dx + dy * dy) - radius));
        }
    s.Commit(dc, rc);
}

void Plot(HDC dc, int x, int y, int w, int h, const unsigned char* coverage, COLORREF colour)
{
    if (!coverage) return;
    const RECT rc{ x, y, x + w, y + h };
    Scratch s;
    if (!s.Open(dc, rc)) return;
    for (int py = 0; py < s.h; ++py)
        for (int px = 0; px < s.w; ++px)
            s.Blend(px, py, colour, coverage[py * w + px] / 255.0);
    s.Commit(dc, rc);
}
}
}
