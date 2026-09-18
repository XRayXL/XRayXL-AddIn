#include "ribbonart.h"

#include <olectl.h>
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
}

IDispatch* DisarmPicture(IDispatch* application, HWND dpiOf)
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
    pic->Release();
    if (got != h) { ReleaseDC(nullptr, dc); source->Release(); return nullptr; }

    // This rendition is pale grey on a clear sheet; the ribbon draws dark ink on a near-white one.
    const auto at = [&](int x, int y) -> Px& { return px[static_cast<size_t>(y) * w + x]; };
    const Px ink{ 0x38, 0x3A, 0x3A, 0xFF }, mark{ 0x74, 0x77, 0x79, 0xFF }, sheet{ 0xFA, 0xFA, 0xFA, 0xFF };

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

    // The dot's solid core is where the square goes.
    int x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (int y = 0; y < h; ++y)
        for (int x = 0; x < w; ++x)
            if (Red(at(x, y), 90) && at(x, y).a >= 0xF0)
            { x0 = min(x0, x); x1 = max(x1, x); y0 = min(y0, y); y1 = max(y1, y); }
    if (x1 < 0) { ReleaseDC(nullptr, dc); source->Release(); return nullptr; }

    for (int y = 0; y < h; ++y)
        for (int x = 0; x < w; ++x)
        {
            Px& p = at(x, y);
            const bool inSquare = x >= x0 && x <= x1 && y >= y0 && y <= y1;
            const Px colour = inSquare ? ink : (p.r < 0xB8 ? mark : ink);     // list marks are the darker grey here
            const int a = inSquare ? 255 : (Red(p, 25) ? 0 : p.a);
            if (outside[static_cast<size_t>(y) * w + x])
            {
                // premultiplied, as a 32-bit bitmap's alpha is read
                p = Px{ static_cast<BYTE>(colour.b * a / 255), static_cast<BYTE>(colour.g * a / 255),
                        static_cast<BYTE>(colour.r * a / 255), static_cast<BYTE>(a) };
                continue;
            }
            p = Px{ static_cast<BYTE>((colour.b * a + sheet.b * (255 - a)) / 255),
                    static_cast<BYTE>((colour.g * a + sheet.g * (255 - a)) / 255),
                    static_cast<BYTE>((colour.r * a + sheet.r * (255 - a)) / 255), 255 };
        }

    ReleaseDC(nullptr, dc);
    source->Release();
    return PictureOf(px, w, h);
}
}
}
}
