#pragma once

#include "softtext.h"

#include <windows.h>

// Excel's own dialog style, free of any one dialog: the palette, the DPI arithmetic, the
// owner-drawn controls, and the subclasses that give the stock ones Excel's frames.

namespace ui
{
namespace excelstyle
{
    using ui::text::Face;

    // ---- palette, matching Excel's own Options dialog -----------------------------------

    constexpr COLORREF kPage      = RGB(0xFF, 0xFF, 0xFF);
    constexpr COLORREF kPane      = RGB(0xF0, 0xF0, 0xF0);
    constexpr COLORREF kItemFace  = RGB(0xFA, 0xFA, 0xFA);   // a selected or hovered category
    constexpr COLORREF kItemHot   = RGB(0xC7, 0xC7, 0xC7);
    constexpr COLORREF kRule      = RGB(0xD1, 0xD1, 0xD1);
    constexpr COLORREF kEdge      = RGB(0x8A, 0x8A, 0x8A);   // every 1px frame at rest
    constexpr COLORREF kEdgeHot   = RGB(0x80, 0x80, 0x80);
    constexpr COLORREF kEdgeOff   = RGB(0xC0, 0xC0, 0xC0);
    constexpr COLORREF kHeavy     = RGB(0x00, 0x00, 0x00);   // every 2px ring
    constexpr COLORREF kHotFace   = RGB(0xF5, 0xF5, 0xF5);
    constexpr COLORREF kDownFace  = RGB(0xE0, 0xE0, 0xE0);
    constexpr COLORREF kText      = RGB(0x24, 0x24, 0x24);
    constexpr COLORREF kDisabled  = RGB(0x9A, 0x9A, 0x9A);
    constexpr COLORREF kCheckOn   = RGB(0x10, 0x7C, 0x41);
    constexpr COLORREF kCheckHot  = RGB(0x0F, 0x70, 0x3B);
    constexpr COLORREF kListFrame = RGB(0x61, 0x61, 0x61);
    constexpr COLORREF kIconInk   = RGB(0x2B, 0x57, 0x9A);

    // ---- the brushes and fonts every control paints with --------------------------------

    // At WM_INITDIALOG and WM_DESTROY. Begin is idempotent within one dialog.
    void Begin(HWND dlg);
    void End();

    // This state is one per process, so only one XRayXL dialog may be open. True, with that
    // dialog brought to the front, if one already is.
    bool ShowOpenDialog();

    HBRUSH PageBrush();
    HBRUSH PaneBrush();
    HBRUSH HotBrush();

    // ClearType for the fields, greyscale for the wrapped texts, as Excel's are.
    void  MakeFonts(HWND dlg);
    HFONT FieldFont();
    HFONT SoftFont();

    // ---- scaling -------------------------------------------------------------------------

    int WindowDpi(HWND h);
    int DpiOf(HWND h);                       // WindowDpi, or XRAYXL_UI_DPI under XRAYXL_DIAG
    int Px(HWND h, int at96);
    int Line(HWND h, int at96);              // hairlines step with the whole-number scale
    int Fit(HWND h, int at96, int at144);    // sizes Excel does not scale linearly

    void Text(HWND h, HDC dc, const RECT& rc, const wchar_t* str, Face face,
              COLORREF colour, unsigned flags);

    // Draws off screen and copies once, so a repaint never shows a half-drawn control.
    template <class Body> void Buffered(HDC dc, const RECT& rc, Body body)
    {
        const int w = rc.right - rc.left, h = rc.bottom - rc.top;
        if (w <= 0 || h <= 0) return;
        HDC mem = CreateCompatibleDC(dc);
        HBITMAP bmp = CreateCompatibleBitmap(dc, w, h);
        if (mem && bmp)
        {
            HGDIOBJ old = SelectObject(mem, bmp);
            body(mem, RECT{ 0, 0, w, h });
            BitBlt(dc, rc.left, rc.top, w, h, mem, 0, 0, SRCCOPY);
            SelectObject(mem, old);
        }
        if (bmp) DeleteObject(bmp);
        if (mem) DeleteDC(mem);
    }

    // ---- the owner-drawn controls ---------------------------------------------------------

    void DrawLabel(const DRAWITEMSTRUCT* di, Face face, unsigned flags);
    void DrawRule(const DRAWITEMSTRUCT* di);
    void DrawButton(const DRAWITEMSTRUCT* di);
    void DrawCheck(const DRAWITEMSTRUCT* di);
    void DrawCategory(const DRAWITEMSTRUCT* di);
    void DrawComboItem(const DRAWITEMSTRUCT* di);
    void DrawPaneFrame(HWND h, HDC dc, const RECT& windowRc);

    // The dotted resize grip a sizable dialog carries in its bottom-right corner, drawn into `corner`.
    void DrawSizeGrip(HWND h, HDC dc, const RECT& corner);

    // The square `corner` wants, at this DPI.
    int SizeGripSide(HWND h);

    // An owner-drawn check box has no state of its own.
    bool IsChecked(HWND dlg, int id);
    void SetChecked(HWND dlg, int id, bool on);

    // ---- subclasses ------------------------------------------------------------------------

    void Subclass(HWND h, WNDPROC with);
    LRESULT Previous(HWND h, UINT m, WPARAM wp, LPARAM lp);

    LRESULT CALLBACK ButtonProc(HWND h, UINT m, WPARAM wp, LPARAM lp);
    LRESULT CALLBACK ComboProc(HWND h, UINT m, WPARAM wp, LPARAM lp);
    LRESULT CALLBACK ListProc(HWND h, UINT m, WPARAM wp, LPARAM lp);
    LRESULT CALLBACK EditProc(HWND h, UINT m, WPARAM wp, LPARAM lp);

    // A text box's own context menu, if the dialog offers one. Called before the stock menu.
    void SetEditMenu(bool (*hook)(HWND edit, LPARAM lp));

    // Whether a text box should paint lit: hovered or focused.
    bool EditLit(HWND h);

    // A drop-down is opening; the popup is shown where it lands rather than slid in.
    void SetDropping(bool on);

    // ---- the system's parts ------------------------------------------------------------------

    void StyleDropList(HWND combo);
    void WhiteCaption(HWND dlg);

    // The module this code is in, for DialogBoxParam and FindResource.
    HMODULE Self();

    // Per-monitor DPI for the length of a dialog: Office calls add-ins system-DPI-aware, which
    // blurs a window on a scaled monitor.
    struct DpiScope
    {
        DpiScope();
        ~DpiScope();
    private:
        HANDLE before = nullptr;
    };
}
}
