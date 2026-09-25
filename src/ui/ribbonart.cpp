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
    IDispatch* GetObjectProp(IDispatch* obj, const wchar_t* name)
    {
        if (!obj) return nullptr;
        DISPID id = 0;
        OLECHAR* n = const_cast<OLECHAR*>(name);
        if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return nullptr;
        DISPPARAMS none{};
        VARIANT out; VariantInit(&out);
        if (FAILED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &none, &out, nullptr, nullptr)))
            return nullptr;
        if (out.vt == VT_DISPATCH) return out.pdispVal;
        VariantClear(&out);
        return nullptr;
    }

    // CommandBars.GetImageMso(name, size, size)
    IDispatch* ImageMso(IDispatch* bars, const wchar_t* name, int size)
    {
        if (!bars) return nullptr;
        DISPID id = 0;
        OLECHAR method[] = L"GetImageMso";
        OLECHAR* n = method;
        if (FAILED(bars->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return nullptr;
        VARIANT args[3];                                   // reverse order
        for (VARIANT& a : args) VariantInit(&a);
        args[0].vt = VT_I4;   args[0].lVal = size;
        args[1].vt = VT_I4;   args[1].lVal = size;
        args[2].vt = VT_BSTR; args[2].bstrVal = SysAllocString(name);
        DISPPARAMS dp{ args, nullptr, 3, 0 };
        VARIANT out; VariantInit(&out);
        const HRESULT hr = bars->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &out, nullptr, nullptr);
        VariantClear(&args[2]);
        if (SUCCEEDED(hr) && out.vt == VT_DISPATCH) return out.pdispVal;
        VariantClear(&out);
        return nullptr;
    }

    int DpiOf(HWND h)
    {
        typedef UINT (WINAPI* Fn)(HWND);
        static const auto fn = reinterpret_cast<Fn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow")));
        const UINT d = (fn && h) ? fn(h) : 0;
        return d ? static_cast<int>(d) : 96;
    }

    struct Px { BYTE b, g, r, a; };
    bool Red(const Px& p, int margin) { return p.r > p.g + margin && p.r > p.b + margin; }

    BITMAPINFO TopDown32(int w, int h)
    {
        BITMAPINFO bi{};
        bi.bmiHeader.biSize = sizeof(bi.bmiHeader);
        bi.bmiHeader.biWidth = w;
        bi.bmiHeader.biHeight = -h;
        bi.bmiHeader.biPlanes = 1;
        bi.bmiHeader.biBitCount = 32;
        return bi;
    }

    // An IPictureDisp over premultiplied pixels; it owns the bitmap.
    IDispatch* PictureOf(const std::vector<Px>& px, int w, int h)
    {
        const BITMAPINFO bi = TopDown32(w, h);
        void* bits = nullptr;
        HDC dc = GetDC(nullptr);
        HBITMAP made = CreateDIBSection(dc, &bi, DIB_RGB_COLORS, &bits, nullptr, 0);
        ReleaseDC(nullptr, dc);
        if (!made || !bits) { if (made) DeleteObject(made); return nullptr; }
        memcpy(bits, px.data(), px.size() * sizeof(Px));

        PICTDESC desc{};
        desc.cbSizeofstruct = sizeof(desc);
        desc.picType = PICTYPE_BITMAP;
        desc.bmp.hbitmap = made;
        IDispatch* out = nullptr;
        if (FAILED(OleCreatePictureIndirect(&desc, IID_IDispatch, TRUE, reinterpret_cast<void**>(&out))) || !out)
        {
            DeleteObject(made);
            return nullptr;
        }
        return out;
    }

    enum class Mark { Record, Stop, FollowEnd };

    // The mark at (du, dv) from the badge's centre, or null for none. `side` is the square's side;
    // the triangle's corners reach as far out as the square's, and it is centred by its centroid.
    const Px* MarkAt(Mark m, double du, double dv, double side, const Px& ink, const Px& green, const Px& red)
    {
        switch (m)
        {
        case Mark::Record:
            return (du * du + dv * dv <= side * side / 4) ? &red : nullptr;
        case Mark::Stop:
            return (std::fabs(du) <= side / 2 && std::fabs(dv) <= side / 2) ? &ink : nullptr;
        case Mark::FollowEnd:
        {
            const double b = side * 1.2247, height = b * 0.8660;        // equilateral, pointing down
            const double top = -height / 3, apex = 2 * height / 3;
            if (dv < top || dv > apex) return nullptr;
            return (std::fabs(du) <= (b / 2) * (apex - dv) / height) ? &green : nullptr;
        }
        }
        return nullptr;
    }

    // Office's MacroRecord pixels, premultiplied, repainted in place: dark ink on a near-white sheet,
    // and its badge redrawn 20% bigger around our mark. False when there is no badge to replace.
    bool Paint(std::vector<Px>& px, int w, int h, Mark m)
    {
        const auto at = [&](int x, int y) -> Px& { return px[static_cast<size_t>(y) * w + x]; };
        const Px ink{ 0x38, 0x3A, 0x3A, 0xFF }, mark{ 0x74, 0x77, 0x79, 0xFF }, sheet{ 0xFA, 0xFA, 0xFA, 0xFF };
        const Px green{ 0x41, 0x7C, 0x10, 0xFF };        // Excel's green, #107C41
        const Px clear{ 0, 0, 0, 0 };

        // Outside the drawing: whatever clear pixels the border can reach.
        std::vector<char> outside(px.size(), 0);
        std::vector<int> todo;
        const auto visit = [&](int x, int y) {
            if (x < 0 || y < 0 || x >= w || y >= h) return;
            const size_t i = static_cast<size_t>(y) * w + x;
            if (outside[i] || px[i].a >= 0x80) return;
            outside[i] = 1; todo.push_back(static_cast<int>(i)); };
        for (int x = 0; x < w; ++x) { visit(x, 0); visit(x, h - 1); }
        for (int y = 0; y < h; ++y) { visit(0, y); visit(w - 1, y); }
        while (!todo.empty())
        {
            const int i = todo.back(); todo.pop_back();
            visit(i % w + 1, i / w); visit(i % w - 1, i / w); visit(i % w, i / w + 1); visit(i % w, i / w - 1);
        }

        // The red dot: its solid core, and all of it with its soft edge.
        RECT core{ w, h, -1, -1 }, dot{ w, h, -1, -1 };
        const auto grow = [](RECT& r, int x, int y) {
            r.left = min(r.left, static_cast<LONG>(x)); r.right  = max(r.right,  static_cast<LONG>(x));
            r.top  = min(r.top,  static_cast<LONG>(y)); r.bottom = max(r.bottom, static_cast<LONG>(y)); };
        for (int y = 0; y < h; ++y)
            for (int x = 0; x < w; ++x)
            {
                if (Red(at(x, y), 90) && at(x, y).a >= 0xF0) grow(core, x, y);
                if (Red(at(x, y), 25) && at(x, y).a > 0)     grow(dot, x, y);
            }
        if (core.right < 0) return false;
        const Px red = at((core.left + core.right) / 2, (core.top + core.bottom) / 2);

        // Office's badge: centred on the dot, its ring the last solid pixel right of it on that row.
        const double cx = (dot.left + dot.right + 1) / 2.0, cy = (dot.top + dot.bottom + 1) / 2.0;
        int edge = -1;
        for (int x = dot.right + 1; x < w; ++x) if (at(x, static_cast<int>(cy)).a >= 0x40) edge = x;
        if (edge < 0) return false;
        const double r0 = edge + 1 - cx;
        // 20% bigger, growing up and left so its right and bottom edges stay where Office's were.
        const double r1 = r0 * 1.2, nx = cx - (r1 - r0), ny = cy - (r1 - r0);
        const double ring = w / 30.0, gap = w / 32.0;
        const double side = (core.right - core.left + 1) * 1.2;

        // The page recoloured, without the old dot.
        std::vector<Px> page(px.size());
        for (int y = 0; y < h; ++y)
            for (int x = 0; x < w; ++x)
            {
                const size_t i = static_cast<size_t>(y) * w + x;
                const Px& p = px[i];
                const Px base = p.r < 0xB8 ? mark : ink;      // list marks are the darker grey here
                const int a = Red(p, 25) ? 0 : p.a;
                // premultiplied outside the drawing, as a 32-bit bitmap's alpha is read; opaque on the sheet
                page[i] = outside[i]
                    ? Px{ static_cast<BYTE>(base.b * a / 255), static_cast<BYTE>(base.g * a / 255),
                          static_cast<BYTE>(base.r * a / 255), static_cast<BYTE>(a) }
                    : Px{ static_cast<BYTE>((base.b * a + sheet.b * (255 - a)) / 255),
                          static_cast<BYTE>((base.g * a + sheet.g * (255 - a)) / 255),
                          static_cast<BYTE>((base.r * a + sheet.r * (255 - a)) / 255), 255 };
            }

        // The new badge over it, sampled 4x4: a ring of ink, the sheet inside it holding the mark, and a
        // clear gap around it, as Office's has. What is left of Office's badge outside it is cleared.
        for (int y = 0; y < h; ++y)
            for (int x = 0; x < w; ++x)
            {
                const size_t i = static_cast<size_t>(y) * w + x;
                const double px0 = x + 0.5, py0 = y + 0.5;
                if (std::hypot(px0 - nx, py0 - ny) > r1 + gap + 1 && std::hypot(px0 - cx, py0 - cy) > r0 + 1)
                {
                    px[i] = page[i];
                    continue;
                }
                int sb = 0, sg = 0, sr = 0, sa = 0;
                for (int sy = 0; sy < 4; ++sy)
                    for (int sx = 0; sx < 4; ++sx)
                    {
                        const double u = x + (sx + 0.5) / 4, v = y + (sy + 0.5) / 4;
                        const double rn = std::hypot(u - nx, v - ny);
                        const Px* s;
                        if (rn <= r1)
                        {
                            const Px* mk = MarkAt(m, u - nx, v - ny, side, ink, green, red);
                            s = rn >= r1 - ring ? &ink : (mk ? mk : &sheet);
                        }
                        else if (rn <= r1 + gap || std::hypot(u - cx, v - cy) <= r0 + 0.5) s = &clear;
                        else s = &page[i];
                        sb += s->b; sg += s->g; sr += s->r; sa += s->a;
                    }
                px[i] = Px{ static_cast<BYTE>(sb / 16), static_cast<BYTE>(sg / 16),
                            static_cast<BYTE>(sr / 16), static_cast<BYTE>(sa / 16) };
            }
        return true;
    }

    IDispatch* RecordPageWith(Mark m, IDispatch* application, HWND dpiOf)
    {
        const int size = MulDiv(32, DpiOf(dpiOf), 96);         // a large button's icon
        IDispatch* bars = GetObjectProp(application, L"CommandBars");
        IDispatch* source = ImageMso(bars, L"MacroRecord", size);
        if (bars) bars->Release();
        if (!source) return nullptr;

        IPicture* pic = nullptr;
        OLE_HANDLE handle = 0;
        if (FAILED(source->QueryInterface(IID_IPicture, reinterpret_cast<void**>(&pic))) || !pic ||
            FAILED(pic->get_Handle(&handle)) || !handle)
        {
            if (pic) pic->Release();
            source->Release();
            return nullptr;
        }
        HBITMAP src = reinterpret_cast<HBITMAP>(static_cast<UINT_PTR>(handle));
        BITMAP bm{};
        GetObjectW(src, sizeof(bm), &bm);
        const int w = bm.bmWidth, h = bm.bmHeight;

        BITMAPINFO bi = TopDown32(w, h);
        std::vector<Px> px(static_cast<size_t>(w) * h);
        HDC dc = GetDC(nullptr);
        const int got = (w > 0 && h > 0) ? GetDIBits(dc, src, 0, h, px.data(), &bi, DIB_RGB_COLORS) : 0;
        ReleaseDC(nullptr, dc);
        pic->Release();
        source->Release();
        if (got != h || !Paint(px, w, h, m)) return nullptr;
        return PictureOf(px, w, h);
    }
}

IDispatch* ArmPicture(IDispatch* application, HWND dpiOf)    { return RecordPageWith(Mark::Record, application, dpiOf); }
IDispatch* DisarmPicture(IDispatch* application, HWND dpiOf) { return RecordPageWith(Mark::Stop, application, dpiOf); }
IDispatch* TailPicture(IDispatch* application, HWND dpiOf)   { return RecordPageWith(Mark::FollowEnd, application, dpiOf); }
}
}
}
