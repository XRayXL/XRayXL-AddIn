#include "softtext.h"

#include <d2d1.h>
#include <dwrite.h>
#include <string>

namespace ui
{
namespace text
{
namespace
{
    typedef HRESULT (WINAPI* D2DCreateFn)(D2D1_FACTORY_TYPE, REFIID, const D2D1_FACTORY_OPTIONS*, void**);
    typedef HRESULT (WINAPI* DWriteCreateFn)(DWRITE_FACTORY_TYPE, REFIID, IUnknown**);

    bool                 g_tried = false;
    ID2D1Factory*        g_d2d   = nullptr;
    IDWriteFactory*      g_dw    = nullptr;
    ID2D1DCRenderTarget* g_rt    = nullptr;
    IDWriteTextFormat*   g_fmt[3] = {};
    int                  g_fmtDpi = 0;

    constexpr int kBrushes = 4;
    ID2D1SolidColorBrush* g_brush[kBrushes] = {};
    COLORREF              g_brushInk[kBrushes] = {};
    int                   g_brushNext = 0;

    template <class T> void Drop(T*& p) { if (p) { p->Release(); p = nullptr; } }

    void DropBrushes()
    {
        for (auto& b : g_brush) Drop(b);
        for (auto& c : g_brushInk) c = 0;
        g_brushNext = 0;
    }

    void DropTarget() { DropBrushes(); Drop(g_rt); }

    // A grid asks for two colours and a dialog for four; anything past that recycles.
    ID2D1SolidColorBrush* BrushFor(COLORREF colour)
    {
        if (!g_rt) return nullptr;
        for (int i = 0; i < kBrushes; ++i)
            if (g_brush[i] && g_brushInk[i] == colour) return g_brush[i];
        const D2D1_COLOR_F c = D2D1::ColorF(GetRValue(colour) / 255.0f, GetGValue(colour) / 255.0f,
                                            GetBValue(colour) / 255.0f);
        ID2D1SolidColorBrush* made = nullptr;
        if (FAILED(g_rt->CreateSolidColorBrush(c, &made))) return nullptr;
        const int at = g_brushNext;
        g_brushNext = (g_brushNext + 1) % kBrushes;
        Drop(g_brush[at]);
        g_brush[at] = made;
        g_brushInk[at] = colour;
        return made;
    }

    // The render target, made once and kept. Software at 96: the rects handed in are pixels.
    bool EnsureTarget()
    {
        if (g_rt) return true;
        const D2D1_RENDER_TARGET_PROPERTIES props = D2D1::RenderTargetProperties(
            D2D1_RENDER_TARGET_TYPE_SOFTWARE,
            D2D1::PixelFormat(DXGI_FORMAT_B8G8R8A8_UNORM, D2D1_ALPHA_MODE_IGNORE), 96.0f, 96.0f);
        if (FAILED(g_d2d->CreateDCRenderTarget(&props, &g_rt))) { g_rt = nullptr; return false; }
        g_rt->SetTextAntialiasMode(D2D1_TEXT_ANTIALIAS_MODE_GRAYSCALE);
        // hinted, as Office's text is: the default mode smears every stem over two columns
        IDWriteRenderingParams* def = nullptr;
        IDWriteRenderingParams* hinted = nullptr;
        if (SUCCEEDED(g_dw->CreateRenderingParams(&def)) && def &&
            SUCCEEDED(g_dw->CreateCustomRenderingParams(def->GetGamma(), def->GetEnhancedContrast(),
                          def->GetClearTypeLevel(), def->GetPixelGeometry(),
                          DWRITE_RENDERING_MODE_GDI_CLASSIC, &hinted)) && hinted)
            g_rt->SetTextRenderingParams(hinted);
        if (hinted) hinted->Release();
        if (def) def->Release();
        return true;
    }

    bool Ensure()
    {
        if (g_tried) return g_d2d && g_dw;
        g_tried = true;
        HMODULE d2d = LoadLibraryExW(L"d2d1.dll",   nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
        HMODULE dw  = LoadLibraryExW(L"dwrite.dll", nullptr, LOAD_LIBRARY_SEARCH_SYSTEM32);
        if (!d2d || !dw) return false;
        auto makeD2D = reinterpret_cast<D2DCreateFn>(GetProcAddress(d2d, "D2D1CreateFactory"));
        auto makeDW  = reinterpret_cast<DWriteCreateFn>(GetProcAddress(dw, "DWriteCreateFactory"));
        if (!makeD2D || !makeDW) return false;
        if (FAILED(makeD2D(D2D1_FACTORY_TYPE_SINGLE_THREADED, __uuidof(ID2D1Factory), nullptr,
                           reinterpret_cast<void**>(&g_d2d)))) g_d2d = nullptr;
        if (FAILED(makeDW(DWRITE_FACTORY_TYPE_SHARED, __uuidof(IDWriteFactory),
                          reinterpret_cast<IUnknown**>(&g_dw)))) g_dw = nullptr;
        return g_d2d && g_dw;
    }

    IDWriteTextFormat* Format(Face face, int dpi)
    {
        if (dpi != g_fmtDpi) { for (auto& f : g_fmt) Drop(f); g_fmtDpi = dpi; }
        IDWriteTextFormat*& f = g_fmt[static_cast<int>(face)];
        if (!f)
        {
            // body and headings are 9pt, the page title 10pt
            const float pt = face == Face::Title ? 10.0f : 9.0f;
            g_dw->CreateTextFormat(L"Segoe UI", nullptr,
                                   face == Face::Bold ? DWRITE_FONT_WEIGHT_BOLD : DWRITE_FONT_WEIGHT_NORMAL,
                                   DWRITE_FONT_STYLE_NORMAL, DWRITE_FONT_STRETCH_NORMAL,
                                   pt * dpi / 72.0f, L"en-us", &f);
        }
        return f;
    }

    // '&x' underlines x; '&&' is a literal ampersand.
    std::wstring StripPrefix(const wchar_t* s, UINT32& underline)
    {
        std::wstring out;
        underline = UINT32_MAX;
        for (; *s; ++s)
        {
            if (*s == L'&' && s[1])
            {
                ++s;
                if (*s != L'&' && underline == UINT32_MAX) underline = static_cast<UINT32>(out.size());
            }
            out.push_back(*s);
        }
        return out;
    }

    IDWriteTextLayout* Layout(const wchar_t* s, Face face, unsigned flags, int dpi, float w, float h)
    {
        IDWriteTextFormat* f = Format(face, dpi);
        if (!f) return nullptr;
        UINT32 underline = UINT32_MAX;
        const std::wstring text = (flags & kPrefix) ? StripPrefix(s, underline) : std::wstring(s);
        IDWriteTextLayout* lay = nullptr;
        // hinted advances too, so every glyph starts on a pixel
        if (FAILED(g_dw->CreateGdiCompatibleTextLayout(text.c_str(), static_cast<UINT32>(text.size()), f, w, h,
                                                       1.0f, nullptr, FALSE, &lay)))
            return nullptr;
        lay->SetWordWrapping((flags & kWrap) ? DWRITE_WORD_WRAPPING_WRAP : DWRITE_WORD_WRAPPING_NO_WRAP);
        lay->SetTextAlignment((flags & kCentre) ? DWRITE_TEXT_ALIGNMENT_CENTER : DWRITE_TEXT_ALIGNMENT_LEADING);
        lay->SetParagraphAlignment((flags & kVCentre) ? DWRITE_PARAGRAPH_ALIGNMENT_CENTER
                                                      : DWRITE_PARAGRAPH_ALIGNMENT_NEAR);
        if (underline != UINT32_MAX) lay->SetUnderline(TRUE, DWRITE_TEXT_RANGE{ underline, 1 });
        if (flags & kUnderline) lay->SetUnderline(TRUE, DWRITE_TEXT_RANGE{ 0, static_cast<UINT32>(text.size()) });
        return lay;
    }
}

bool Ready() { return Ensure(); }

bool Draw(HDC dc, const RECT& rc, const wchar_t* s, Face face, COLORREF colour, unsigned flags, int dpi, float dx)
{
    if (!s || !Ensure() || rc.right <= rc.left || rc.bottom <= rc.top) return false;
    if (!EnsureTarget()) return false;
    IDWriteTextLayout* lay = Layout(s, face, flags, dpi, static_cast<float>(rc.right - rc.left),
                                    static_cast<float>(rc.bottom - rc.top));
    if (!lay) return false;

    bool ok = false;
    if (SUCCEEDED(g_rt->BindDC(dc, &rc)))
    {
        g_rt->BeginDraw();
        g_rt->SetTransform(D2D1::Matrix3x2F::Identity());
        if (ID2D1SolidColorBrush* ink = BrushFor(colour))
        {
            g_rt->DrawTextLayout(D2D1::Point2F(dx, 0.0f), lay, ink);
            ok = true;
        }
        if (FAILED(g_rt->EndDraw())) { DropTarget(); ok = false; }
    }
    lay->Release();
    return ok;
}

Batch::Batch(HDC dc, const RECT& area)
    : m_area(area)
{
    if (!Ensure() || !EnsureTarget()) return;
    if (FAILED(g_rt->BindDC(dc, &area))) return;
    g_rt->BeginDraw();
    g_rt->SetTransform(D2D1::Matrix3x2F::Identity());
    m_ok = true;
}

Batch::~Batch()
{
    if (!m_ok || !g_rt) return;
    if (FAILED(g_rt->EndDraw())) DropTarget();
}

bool Batch::Put(const RECT& rc, const RECT& clip, const wchar_t* s, Face face,
                COLORREF colour, unsigned flags, int dpi)
{
    if (!m_ok || !g_rt || !s || rc.right <= rc.left || rc.bottom <= rc.top) return false;
    if (clip.right <= clip.left || clip.bottom <= clip.top) return false;
    IDWriteTextLayout* lay = Layout(s, face, flags, dpi, static_cast<float>(rc.right - rc.left),
                                    static_cast<float>(rc.bottom - rc.top));
    if (!lay) return false;

    ID2D1SolidColorBrush* ink = BrushFor(colour);
    if (ink)
    {
        const D2D1_RECT_F box = D2D1::RectF(
            static_cast<float>(clip.left   - m_area.left), static_cast<float>(clip.top    - m_area.top),
            static_cast<float>(clip.right  - m_area.left), static_cast<float>(clip.bottom - m_area.top));
        g_rt->PushAxisAlignedClip(box, D2D1_ANTIALIAS_MODE_ALIASED);
        g_rt->DrawTextLayout(D2D1::Point2F(static_cast<float>(rc.left - m_area.left),
                                           static_cast<float>(rc.top  - m_area.top)), lay, ink);
        g_rt->PopAxisAlignedClip();
    }
    lay->Release();
    return ink != nullptr;
}

bool Measure(const wchar_t* s, Face face, unsigned flags, int dpi, SIZE& out)
{
    if (!s || !Ensure()) return false;
    IDWriteTextLayout* lay = Layout(s, face, flags & ~(kCentre | kVCentre | kWrap), dpi, 4096.0f, 4096.0f);
    if (!lay) return false;
    DWRITE_TEXT_METRICS m{};
    const bool ok = SUCCEEDED(lay->GetMetrics(&m));
    lay->Release();
    if (!ok) return false;
    out.cx = static_cast<LONG>(m.widthIncludingTrailingWhitespace + 0.999f);
    out.cy = static_cast<LONG>(m.height + 0.999f);
    return true;
}

void Release()
{
    DropTarget();
    for (auto& f : g_fmt) Drop(f);
    g_fmtDpi = 0;
    Drop(g_dw);
    Drop(g_d2d);
    g_tried = false;
}
}
}
