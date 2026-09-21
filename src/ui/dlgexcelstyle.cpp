#include "dlgexcelstyle.h"
#include "softdraw.h"

#include "core/tracemodes.h"

#include <algorithm>
#include <cstdlib>

namespace ui
{
namespace excelstyle
{
namespace
{
    HBRUSH g_white = nullptr, g_pane = nullptr, g_hot = nullptr;
    HFONT  g_edit = nullptr;        // text boxes draw their own text: ClearType, as Excel's
    HFONT  g_editSoft = nullptr;    // the wrapped texts: greyscale, as Excel's
    int    g_hotItem = -1;          // the category under the pointer
    bool   g_dropping = false;      // a drop-down is opening
    HWND   g_hotCtl  = nullptr;     // the control under the pointer

    bool (*g_editMenu)(HWND, LPARAM) = nullptr;

    const wchar_t* const kOldProc = L"XRayOldProc";

    HFONT EditFont(HWND dlg, DWORD quality)
    {
        return CreateFontW(-MulDiv(9, DpiOf(dlg), 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                           DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, quality,
                           DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    }

    bool IsMultiline(HWND h) { return (GetWindowLongPtrW(h, GWL_STYLE) & ES_MULTILINE) != 0; }

    bool DwmSet(HWND h, DWORD attr, const void* v, DWORD cb)
    {
        HMODULE dwm = LoadLibraryW(L"dwmapi.dll");
        if (!dwm) return false;
        using Fn = HRESULT (WINAPI*)(HWND, DWORD, LPCVOID, DWORD);
        auto set = reinterpret_cast<Fn>(reinterpret_cast<void*>(GetProcAddress(dwm, "DwmSetWindowAttribute")));
        const bool ok = set && SUCCEEDED(set(h, attr, v, cb));
        FreeLibrary(dwm);
        return ok;
    }

    // Excel's checked box at 96 DPI: a one-pixel white line, 43% either side, soft corners.
    void PlotTick(HDC dc, const RECT& b)
    {
        unsigned char m[13 * 13] = {};
        auto at = [&m](int x, int y) -> unsigned char& { return m[y * 13 + x]; };
        const POINT tick[] = { {3,6}, {4,7}, {5,8}, {6,7}, {7,6}, {8,5}, {9,4} };
        for (const POINT& p : tick)
            for (int dx = -1; dx <= 1; dx += 2)
            {
                unsigned char& c = at(p.x + dx, p.y);
                c = static_cast<unsigned char>(255 - (255 - c) * (255 - 110) / 255);
            }
        for (const POINT& p : tick) at(p.x, p.y) = 255;
        for (int cy = 0; cy <= 12; cy += 12)
            for (int cx = 0; cx <= 12; cx += 12)
            {
                at(cx, cy) = 158;
                at(cx ? cx - 1 : 1, cy) = 54;
                at(cx, cy ? cy - 1 : 1) = 54;
            }
        ui::soft::Plot(dc, b.left, b.top, 13, 13, m, kPage);
    }

    void PaintCombo(HWND h, HDC target)
    {
        RECT client;
        GetClientRect(h, &client);
        const bool disabled = !IsWindowEnabled(h);
        const bool hot      = (h == g_hotCtl) && !disabled;
        const int sel = static_cast<int>(SendMessageW(h, CB_GETCURSEL, 0, 0));
        wchar_t text[64] = {};
        if (sel >= 0) SendMessageW(h, CB_GETLBTEXT, sel, reinterpret_cast<LPARAM>(text));

        Buffered(target, client, [&](HDC dc, const RECT& r) {
            FillRect(dc, &r, g_white);
            RECT btn = r;
            btn.left = r.right - Px(h, 15);
            const COLORREF edge = disabled ? kEdgeOff : kEdge;
            ui::soft::RoundRect(dc, r, ui::soft::Box{ kPage, edge, Line(h, 1), Px(h, 4), true });
            if (hot)
            {
                // the button end goes grey behind a divider, keeping the frame's curve
                HRGN clip = CreateRectRgn(btn.left, r.top, r.right, r.bottom);
                SelectClipRgn(dc, clip);
                ui::soft::RoundRect(dc, r, ui::soft::Box{ kHotFace, edge, Line(h, 1), Px(h, 4), true });
                SelectClipRgn(dc, nullptr);
                DeleteObject(clip);
                const RECT divider{ btn.left, r.top + Line(h, 1), btn.left + Line(h, 1), r.bottom - Line(h, 1) };
                HBRUSH b = CreateSolidBrush(edge);
                FillRect(dc, &divider, b);
                DeleteObject(b);
            }
            const int cx = (btn.left + btn.right) / 2, cy = r.bottom / 2, half = Px(h, 3);
            const POINT v[3] = { { cx - half, cy - half }, { cx, cy }, { cx + half, cy - half } };
            ui::soft::Stroke(dc, v, 3, Px(h, 1), disabled ? kDisabled : kText);

            RECT t = r;
            t.left += Px(h, 6);
            t.right = btn.left - Px(h, 2);
            Text(h, dc, t, text, Face::Body, disabled ? kDisabled : kText, ui::text::kVCentre); });
    }

    // Owner-drawn controls are never told the pointer is over them.
    bool TrackHover(HWND h, UINT m)
    {
        if (m == WM_MOUSEMOVE && g_hotCtl != h)
        {
            HWND was = g_hotCtl;
            g_hotCtl = h;
            if (was) RedrawWindow(was, nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE);
            TRACKMOUSEEVENT t{ sizeof(t), TME_LEAVE, h, 0 };
            TrackMouseEvent(&t);
            return true;
        }
        if ((m == WM_MOUSELEAVE || m == WM_NCDESTROY) && g_hotCtl == h) { g_hotCtl = nullptr; return true; }
        return m == WM_MOUSELEAVE;
    }

    void InvalidateItem(HWND list, int item)
    {
        RECT r;
        if (item >= 0 && SendMessageW(list, LB_GETITEMRECT, item, reinterpret_cast<LPARAM>(&r)) != LB_ERR)
            InvalidateRect(list, &r, FALSE);
    }

    void PaintEditFrame(HWND h)
    {
        RECT w, c;
        GetWindowRect(h, &w);
        GetClientRect(h, &c);
        MapWindowPoints(h, nullptr, reinterpret_cast<POINT*>(&c), 2);
        HDC dc = GetWindowDC(h);
        ExcludeClipRect(dc, c.left - w.left, c.top - w.top, c.right - w.left, c.bottom - w.top);
        SCROLLBARINFO sb{};
        sb.cbSize = sizeof(sb);
        if (GetScrollBarInfo(h, OBJID_VSCROLL, &sb) && !(sb.rgstate[0] & STATE_SYSTEM_INVISIBLE))
            ExcludeClipRect(dc, sb.rcScrollBar.left - w.left, sb.rcScrollBar.top - w.top,
                            sb.rcScrollBar.right - w.left, sb.rcScrollBar.bottom - w.top);
        OffsetRect(&w, -w.left, -w.top);
        FillRect(dc, &w, g_white);
        const bool disabled = !IsWindowEnabled(h);
        const bool lit      = EditLit(h);
        const bool focused  = lit && GetFocus() == h;
        ui::soft::RoundRect(dc, w, ui::soft::Box{
            lit ? kHotFace : kPage,
            disabled ? kEdgeOff : (focused ? kHeavy : (lit ? kEdgeHot : kEdge)),
            Line(h, focused ? 2 : 1), Px(h, 4), true });
        ReleaseDC(h, dc);
    }

    // The popup's frame is DWM's: it follows the rounded window, which a painted ring cannot.
    LRESULT CALLBACK ListFrameProc(HWND h, UINT m, WPARAM wp, LPARAM lp)
    {
        // shown as it is positioned, so AnimateWindow has nothing to slide
        if (m == WM_WINDOWPOSCHANGING && g_dropping)
        {
            WINDOWPOS* pos = reinterpret_cast<WINDOWPOS*>(lp);
            if (!(pos->flags & SWP_HIDEWINDOW)) pos->flags |= SWP_SHOWWINDOW;
        }
        if (m == WM_NCPAINT)
        {
            RECT w, c;
            GetWindowRect(h, &w);
            GetClientRect(h, &c);
            MapWindowPoints(h, nullptr, reinterpret_cast<POINT*>(&c), 2);
            if (HDC dc = GetWindowDC(h))
            {
                ExcludeClipRect(dc, c.left - w.left, c.top - w.top, c.right - w.left, c.bottom - w.top);
                OffsetRect(&w, -w.left, -w.top);
                FillRect(dc, &w, g_white);
                ReleaseDC(h, dc);
            }
            return 0;
        }
        return Previous(h, m, wp, lp);
    }
}

// ---- brushes and fonts -----------------------------------------------------------------

void Begin()
{
    g_hotItem = -1;
    g_hotCtl = nullptr;
    g_dropping = false;
    if (!g_white) g_white = CreateSolidBrush(kPage);
    if (!g_pane)  g_pane  = CreateSolidBrush(kPane);
    if (!g_hot)   g_hot   = CreateSolidBrush(kHotFace);
}

void End()
{
    for (HBRUSH* b : { &g_white, &g_pane, &g_hot }) if (*b) { DeleteObject(*b); *b = nullptr; }
    for (HFONT* f : { &g_edit, &g_editSoft })       if (*f) { DeleteObject(*f); *f = nullptr; }
    g_editMenu = nullptr;
    ui::text::Release();
}

HBRUSH PageBrush() { return g_white; }
HBRUSH PaneBrush() { return g_pane; }
HBRUSH HotBrush()  { return g_hot; }

void MakeFonts(HWND dlg)
{
    if (g_edit)     DeleteObject(g_edit);
    if (g_editSoft) DeleteObject(g_editSoft);
    g_edit     = EditFont(dlg, DEFAULT_QUALITY);
    g_editSoft = EditFont(dlg, ANTIALIASED_QUALITY);
}

HFONT FieldFont() { return g_edit; }
HFONT SoftFont()  { return g_editSoft; }

// ---- scaling ---------------------------------------------------------------------------

int WindowDpi(HWND h)
{
    static auto fn = reinterpret_cast<UINT (WINAPI*)(HWND)>(reinterpret_cast<void*>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"), "GetDpiForWindow")));
    if (fn && h) { const UINT d = fn(h); if (d) return static_cast<int>(d); }
    HDC dc = GetDC(nullptr);
    const int d = dc ? GetDeviceCaps(dc, LOGPIXELSX) : 96;
    if (dc) ReleaseDC(nullptr, dc);
    return d ? d : 96;
}

int DpiOf(HWND h)
{
    // XRAYXL_DIAG only: XRAYXL_UI_DPI lays the dialog out as if at that DPI.
    static const int forced = [] {
        wchar_t v[8] = {};
        if (!core::modes::DiagEnabled() || !GetEnvironmentVariableW(L"XRAYXL_UI_DPI", v, 8)) return 0;
        const int d = _wtoi(v);
        return (d >= 96 && d <= 480) ? d : 0; }();
    return forced ? forced : WindowDpi(h);
}

int Px(HWND h, int at96) { return MulDiv(at96, DpiOf(h), 96); }
int Line(HWND h, int at96) { return at96 * (std::max)(1, DpiOf(h) / 96); }
int Fit(HWND h, int at96, int at144) { return at96 + MulDiv(at144 - at96, DpiOf(h) - 96, 48); }

void Text(HWND h, HDC dc, const RECT& rc, const wchar_t* str, Face face, COLORREF colour, unsigned flags)
{
    ui::text::Draw(dc, rc, str, face, colour, flags, DpiOf(h));
}

// ---- the owner-drawn controls -------------------------------------------------------------

void DrawLabel(const DRAWITEMSTRUCT* di, Face face, unsigned flags)
{
    wchar_t text[256] = {};
    GetWindowTextW(di->hwndItem, text, 256);
    Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
        FillRect(dc, &rc, g_white);
        Text(di->hwndItem, dc, rc, text, face, kText, flags); });
}

void DrawRule(const DRAWITEMSTRUCT* di)
{
    HBRUSH rule = CreateSolidBrush(kRule);
    FillRect(di->hDC, &di->rcItem, rule);
    DeleteObject(rule);
}

void DrawButton(const DRAWITEMSTRUCT* di)
{
    HWND h = di->hwndItem;
    const bool disabled = (di->itemState & ODS_DISABLED) != 0;
    const bool pressed  = (di->itemState & ODS_SELECTED) != 0;
    const bool hot      = (h == g_hotCtl) && !disabled;
    // ODS_DEFAULT never arrives for an owner-drawn button, so the dialog is asked.
    const LRESULT defid = SendMessageW(GetParent(h), DM_GETDEFID, 0, 0);
    const bool isDefault = HIWORD(defid) == DC_HASDEFID && LOWORD(defid) == static_cast<WORD>(di->CtlID);
    const bool heavy = hot || pressed || isDefault || (di->itemState & ODS_FOCUS) != 0;
    wchar_t text[64] = {};
    GetWindowTextW(h, text, 64);

    Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
        FillRect(dc, &rc, g_white);
        ui::soft::RoundRect(dc, rc, ui::soft::Box{
            pressed ? kDownFace : (hot ? kHotFace : kPage),
            disabled ? kEdgeOff : (heavy ? kHeavy : kEdge),
            Line(h, heavy ? 2 : 1), Px(h, 4), true });
        RECT t = rc;
        t.top += Px(h, 2);                 // sits one pixel below centre, as Excel's does
        Text(h, dc, t, text, Face::Body, disabled ? kDisabled : kText,
             ui::text::kCentre | ui::text::kVCentre | ui::text::kPrefix); });
}

void DrawCheck(const DRAWITEMSTRUCT* di)
{
    HWND h = di->hwndItem;
    const bool on       = GetWindowLongPtrW(h, GWLP_USERDATA) != 0;
    const bool disabled = (di->itemState & ODS_DISABLED) != 0;
    const bool focused  = (di->itemState & ODS_FOCUS) != 0;
    const bool hot      = (h == g_hotCtl) && !disabled;
    wchar_t text[128] = {};
    GetWindowTextW(h, text, 128);

    Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
        FillRect(dc, &rc, g_white);
        const int side = Px(h, 13);
        const RECT b{ 0, (rc.bottom - side) / 2, side, (rc.bottom - side) / 2 + side };

        COLORREF edge, fill;
        if (disabled) { edge = kEdgeOff; fill = kPage; }
        else if (on)  { edge = fill = hot ? kCheckHot : kCheckOn; }
        else          { edge = hot ? kEdgeHot : kEdge; fill = hot ? kHotFace : kPage; }

        if (on && !disabled && side == 13)
        {
            HBRUSH green = CreateSolidBrush(fill);
            FillRect(dc, &b, green);
            DeleteObject(green);
            PlotTick(dc, b);
        }
        else
        {
            // the border scales with the DPI here, unlike the other frames
            ui::soft::RoundRect(dc, b, ui::soft::Box{ fill, edge, on ? 0 : Px(h, 1), Px(h, 2), true });
            if (on)
            {
                const POINT v[3] = { { side * 3 / 13, b.top + side * 6 / 13 },
                                     { side * 5 / 13, b.top + side * 8 / 13 },
                                     { side * 9 / 13, b.top + side * 4 / 13 } };
                ui::soft::Stroke(dc, v, 3, Px(h, 14) / 10.0, disabled ? RGB(0x90, 0x90, 0x90) : kPage);
            }
        }

        RECT t = rc;
        t.left = b.right + Px(h, 5);
        t.top += Px(h, 2);                 // sits one pixel below centre, as Excel's does
        Text(h, dc, t, text, Face::Body, disabled ? kDisabled : kText, ui::text::kVCentre | ui::text::kPrefix);

        if (!focused) return;
        SIZE sz{};
        ui::text::Measure(text, Face::Body, ui::text::kPrefix, DpiOf(h), sz);
        RECT f{ t.left - Px(h, 3), 0, t.left + sz.cx + Px(h, 3), 0 };
        f.top    = (rc.bottom - sz.cy) / 2 - Px(h, 2);
        f.bottom = f.top + sz.cy + 2 * Px(h, 2);
        HPEN pen = CreatePen(PS_INSIDEFRAME, Line(h, 2), kHeavy);
        HGDIOBJ op = SelectObject(dc, pen);
        HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_BRUSH));
        Rectangle(dc, f.left, f.top, f.right, f.bottom);
        SelectObject(dc, ob);
        SelectObject(dc, op);
        DeleteObject(pen); });
}

// The pane's frame curves into the client corners, so rows and the space below redraw it.
void DrawPaneFrame(HWND h, HDC dc, const RECT& windowRc)
{
    ui::soft::RoundRect(dc, windowRc, ui::soft::Box{ 0, kEdge, Line(h, 1), Px(h, 4), false });
}

// The dotted grip a sizable dialog carries in its bottom-right corner: six dots in a triangle.
// The classic grip embosses each with a white highlight, which draws nothing on a white page.
void DrawSizeGrip(HWND h, HDC dc, const RECT& corner)
{
    const int dot = Px(h, 2), pitch = Px(h, 4), inset = Px(h, 3);
    HBRUSH ink = CreateSolidBrush(RGB(0xA0, 0xA0, 0xA0));
    for (int dx = 0; dx <= 2; ++dx)
        for (int dy = 0; dy + dx <= 2; ++dy)
        {
            const int x = corner.right - inset - dot - dx * pitch;
            const int y = corner.bottom - inset - dot - dy * pitch;
            if (x < corner.left || y < corner.top) continue;
            RECT face{ x, y, x + dot, y + dot };
            FillRect(dc, &face, ink);
        }
    DeleteObject(ink);
}

int SizeGripSide(HWND h) { return Px(h, 14); }

void DrawCategory(const DRAWITEMSTRUCT* di)
{
    if (di->itemID == static_cast<UINT>(-1)) return;
    HWND h = di->hwndItem;
    const bool selected = (di->itemState & ODS_SELECTED) != 0;
    const bool hot      = (static_cast<int>(di->itemID) == g_hotItem) && !selected;
    wchar_t text[64] = {};
    SendMessageW(h, LB_GETTEXT, di->itemID, reinterpret_cast<LPARAM>(text));
    RECT pane;
    GetClientRect(h, &pane);
    InflateRect(&pane, Line(h, 1), Line(h, 1));
    OffsetRect(&pane, -di->rcItem.left, -di->rcItem.top);

    Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
        FillRect(dc, &rc, g_pane);
        if (selected || hot)
            ui::soft::RoundRect(dc, rc, ui::soft::Box{ kItemFace, selected ? kHeavy : kItemHot,
                                                       Line(h, selected ? 2 : 1), Px(h, 4), true });
        DrawPaneFrame(h, dc, pane);
        RECT t = rc;
        t.left += Px(h, 13);
        Text(h, dc, t, text, Face::Body, kText, ui::text::kVCentre); });
}

// A row of the open list; ODS_SELECTED there means "under the pointer".
void DrawComboItem(const DRAWITEMSTRUCT* di)
{
    if (di->itemState & ODS_COMBOBOXEDIT) return;       // the closed field is PaintCombo's
    const bool current = (di->itemState & ODS_SELECTED) != 0;
    wchar_t text[64] = {};
    if (di->itemID != static_cast<UINT>(-1))
        SendMessageW(di->hwndItem, CB_GETLBTEXT, di->itemID, reinterpret_cast<LPARAM>(text));
    Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
        FillRect(dc, &rc, current ? g_hot : g_white);
        RECT t = rc;
        t.left += Px(di->hwndItem, 6);
        Text(di->hwndItem, dc, t, text, Face::Body, kText, ui::text::kVCentre); });
}

bool IsChecked(HWND dlg, int id)
{
    HWND h = GetDlgItem(dlg, id);
    return h && GetWindowLongPtrW(h, GWLP_USERDATA) != 0;
}

void SetChecked(HWND dlg, int id, bool on)
{
    HWND h = GetDlgItem(dlg, id);
    if (!h) return;
    SetWindowLongPtrW(h, GWLP_USERDATA, on ? 1 : 0);
    InvalidateRect(h, nullptr, TRUE);
}

// ---- subclasses -----------------------------------------------------------------------------

void Subclass(HWND h, WNDPROC with)
{
    if (!h) return;
    auto old = reinterpret_cast<WNDPROC>(SetWindowLongPtrW(h, GWLP_WNDPROC, reinterpret_cast<LONG_PTR>(with)));
    SetPropW(h, kOldProc, reinterpret_cast<HANDLE>(old));
}

LRESULT Previous(HWND h, UINT m, WPARAM wp, LPARAM lp)
{
    auto prev = reinterpret_cast<WNDPROC>(GetPropW(h, kOldProc));
    if (m == WM_NCDESTROY) RemovePropW(h, kOldProc);
    return prev ? CallWindowProcW(prev, h, m, wp, lp) : DefWindowProcW(h, m, wp, lp);
}

LRESULT CALLBACK ButtonProc(HWND h, UINT m, WPARAM wp, LPARAM lp)
{
    if (TrackHover(h, m)) InvalidateRect(h, nullptr, FALSE);
    return Previous(h, m, wp, lp);
}

LRESULT CALLBACK ComboProc(HWND h, UINT m, WPARAM wp, LPARAM lp)
{
    if (m == WM_ERASEBKGND) return 1;
    if (m == WM_PAINT)
    {
        PAINTSTRUCT ps;
        PaintCombo(h, BeginPaint(h, &ps));
        EndPaint(h, &ps);
        return 0;
    }
    if (TrackHover(h, m)) InvalidateRect(h, nullptr, FALSE);
    return Previous(h, m, wp, lp);
}

LRESULT CALLBACK ListProc(HWND h, UINT m, WPARAM wp, LPARAM lp)
{
    switch (m)
    {
    case WM_NCCALCSIZE:
        InflateRect(reinterpret_cast<RECT*>(lp), -Line(h, 1), -Line(h, 1));
        return 0;
    case WM_NCPAINT:
    {
        RECT w;
        GetWindowRect(h, &w);
        OffsetRect(&w, -w.left, -w.top);
        const int e = Line(h, 1);
        HDC dc = GetWindowDC(h);
        ExcludeClipRect(dc, e, e, w.right - e, w.bottom - e);
        FillRect(dc, &w, g_white);
        DrawPaneFrame(h, dc, w);
        ReleaseDC(h, dc);
        return 0;
    }
    case WM_ERASEBKGND:
        return 1;                       // rows paint themselves; WM_PAINT fills below them
    case WM_PAINT:
    {
        // The rows paint themselves; what is left is the space under the last one and the
        // frame, drawn over the whole client so no corner depends on what was invalidated.
        const LRESULT res = Previous(h, m, wp, lp);
        RECT c, last{};
        GetClientRect(h, &c);
        const LRESULT n = SendMessageW(h, LB_GETCOUNT, 0, 0);
        if (n > 0) SendMessageW(h, LB_GETITEMRECT, n - 1, reinterpret_cast<LPARAM>(&last));
        if (HDC dc = GetDC(h))
        {
            if (last.bottom < c.bottom)
            {
                RECT below = c;
                below.top = last.bottom;
                FillRect(dc, &below, g_pane);
            }
            RECT frame = c;
            InflateRect(&frame, Line(h, 1), Line(h, 1));
            DrawPaneFrame(h, dc, frame);
            ReleaseDC(h, dc);
        }
        return res;
    }
    case WM_SIZE:
        // The frame's corners move with the control, so the whole of it has to be redrawn:
        // only the newly exposed strip is invalidated otherwise, and the old corners stay.
        RedrawWindow(h, nullptr, nullptr, RDW_INVALIDATE | RDW_FRAME | RDW_ERASE);
        break;
    case WM_MOUSEMOVE:
    {
        const LRESULT r = SendMessageW(h, LB_ITEMFROMPOINT, 0, lp);
        const int item = (HIWORD(r) == 0) ? LOWORD(r) : -1;
        if (item != g_hotItem)
        {
            // only the two rows that changed, or the selected row flickers
            InvalidateItem(h, g_hotItem);
            InvalidateItem(h, item);
            g_hotItem = item;
            TRACKMOUSEEVENT t{ sizeof(t), TME_LEAVE, h, 0 };
            TrackMouseEvent(&t);
        }
        break;
    }
    case WM_MOUSELEAVE:
        InvalidateItem(h, g_hotItem);
        g_hotItem = -1;
        break;
    }
    return Previous(h, m, wp, lp);
}

bool EditLit(HWND h)
{
    return IsWindowEnabled(h) && !IsMultiline(h) && (g_hotCtl == h || GetFocus() == h);
}

void SetEditMenu(bool (*hook)(HWND, LPARAM)) { g_editMenu = hook; }

// Text boxes: the stock edit edits, and Excel's frame is painted in a margin around it.
LRESULT CALLBACK EditProc(HWND h, UINT m, WPARAM wp, LPARAM lp)
{
    const bool multi = IsMultiline(h);
    switch (m)
    {
    case WM_NCCALCSIZE:
    {
        RECT* r = reinterpret_cast<RECT*>(lp);
        if (multi)
        {
            // the margin first, so the scroll bar lands inside the frame
            InflateRect(r, -Px(h, 6), -Px(h, 4));
            r->left += Px(h, 1);
            return Previous(h, m, wp, lp);
        }
        int textH = Px(h, 15);
        HDC dc = GetDC(h);
        HFONT f = reinterpret_cast<HFONT>(SendMessageW(h, WM_GETFONT, 0, 0));
        HGDIOBJ of = f ? SelectObject(dc, f) : nullptr;
        TEXTMETRICW tm{};
        if (GetTextMetricsW(dc, &tm)) textH = tm.tmHeight;
        if (of) SelectObject(dc, of);
        ReleaseDC(h, dc);
        r->left  += Px(h, 7);
        r->right -= Px(h, 7);
        r->top   += ((r->bottom - r->top) - textH) / 2;
        r->bottom = r->top + textH;
        return 0;
    }
    case WM_NCHITTEST:
        if (!multi) return HTCLIENT;     // the margin is part of the box to the mouse
        break;
    case WM_CONTEXTMENU:
        if (g_editMenu && g_editMenu(h, lp)) return 0;
        break;
    case WM_NCPAINT:
        if (multi) Previous(h, m, wp, lp);      // the scroll bar
        PaintEditFrame(h);
        return 0;
    case WM_SETFOCUS:
    case WM_KILLFOCUS:
    case WM_ENABLE:
        RedrawWindow(h, nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE);
        break;
    }
    if (!multi && TrackHover(h, m)) RedrawWindow(h, nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_ERASE);
    return Previous(h, m, wp, lp);
}

void SetDropping(bool on) { g_dropping = on; }

// ---- the system's parts ---------------------------------------------------------------------

void StyleDropList(HWND combo)
{
    COMBOBOXINFO ci{};
    ci.cbSize = sizeof(ci);
    if (!combo || !GetComboBoxInfo(combo, &ci) || !ci.hwndList) return;
    const DWORD roundSmall = 3;                         // DWMWCP_ROUNDSMALL
    DwmSet(ci.hwndList, 33 /*DWMWA_WINDOW_CORNER_PREFERENCE*/, &roundSmall, sizeof(roundSmall));
    DwmSet(ci.hwndList, 34 /*DWMWA_BORDER_COLOR*/, &kListFrame, sizeof(kListFrame));
    if (!GetPropW(ci.hwndList, kOldProc)) Subclass(ci.hwndList, ListFrameProc);
}

void WhiteCaption(HWND dlg)
{
    DwmSet(dlg, 35 /*DWMWA_CAPTION_COLOR*/, &kPage, sizeof(kPage));
    DwmSet(dlg, 36 /*DWMWA_TEXT_COLOR*/,    &kText, sizeof(kText));
}

HMODULE Self()
{
    HMODULE self = nullptr;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                       reinterpret_cast<LPCWSTR>(&Self), &self);
    return self;
}

DpiScope::DpiScope()
{
    typedef HANDLE (WINAPI* SetContextFn)(HANDLE);
    const auto setContext = reinterpret_cast<SetContextFn>(reinterpret_cast<void*>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetThreadDpiAwarenessContext")));
    const HANDLE perMonitorV2 = reinterpret_cast<HANDLE>(static_cast<INT_PTR>(-4));
    before = setContext ? setContext(perMonitorV2) : nullptr;
}

DpiScope::~DpiScope()
{
    if (!before) return;
    typedef HANDLE (WINAPI* SetContextFn)(HANDLE);
    const auto setContext = reinterpret_cast<SetContextFn>(reinterpret_cast<void*>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetThreadDpiAwarenessContext")));
    if (setContext) setContext(before);
}
}
}
