#include "optionsdlg.h"
#include "optionsres.h"
#include "ribbonmodel.h"
#include "traceactions.h"
#include "softdraw.h"
#include "softtext.h"
#include "glyphs.h"
#include "reflow.h"

#include <algorithm>

#include "app/session.h"
#include "core/log.h"
#include "core/tracemodes.h"
#include "core/excel_api.h"
#include "core/text.h"
#include "emit/csv.h"

#include <windows.h>
#include <cwchar>
#include <string>

// Excel's options dialog is owner drawn, we do our best to match its look and feel.

namespace ui
{
namespace options
{
namespace
{
    namespace M = ui::ribbon::model;
    using ui::text::Face;

    // ---- pages ---------------------------------------------------------------

    struct Page { const wchar_t* name; const int* ids; int count; };

    const int kCapture[] = { IDC_CAP_HDR, IDC_XLL_SEC1, IDC_XLL_RULE1, IDC_XLL_DEPTHLBL, IDC_XLL_DEPTH,
                             IDC_XLL_ARGS, IDC_XLL_RET,
                             IDC_VBA_SEC1, IDC_VBA_RULE1, IDC_VBA_DEPTHLBL, IDC_VBA_DEPTH,
                             IDC_VBA_ARGS, IDC_VBA_RET, IDC_VBA_OBJ };
    const int kOut[]     = { IDC_OUT_HDR, IDC_OUT_SEC1, IDC_OUT_RULE1, IDC_OUT_FMTLBL, IDC_OUT_FMT,
                             IDC_OUT_DIRLBL, IDC_OUT_DIR, IDC_OUT_FILELBL, IDC_OUT_FILE, IDC_OUT_TAIL };
    const int kAdvanced[] = { IDC_ADV_HDR, IDC_ADV_SEC1, IDC_ADV_RULE1, IDC_ADV_BUFLBL, IDC_ADV_BUF,
                              IDC_ADV_BUFHINT, IDC_ADV_FULLLBL, IDC_ADV_FULL,
                              IDC_ADV_SEC2, IDC_ADV_RULE2, IDC_ADV_LVLLBL, IDC_ADV_LVL, IDC_ADV_LVLNOTE,
                              IDC_ADV_LOGLBL, IDC_ADV_LOG };
    const int kAbout[]   = { IDC_ABT_HDR, IDC_ABT_OWNER, IDC_ABT_LICLBL, IDC_ABT_LICENSE };
    const int kNotices[] = { IDC_NOT_HDR, IDC_NOT_TEXT };

#define XRAY_PAGE(name, ids) { name, ids, static_cast<int>(sizeof(ids) / sizeof(int)) }
    const Page kPages[] = {
        XRAY_PAGE(L"Capture", kCapture), XRAY_PAGE(L"Output", kOut),
        XRAY_PAGE(L"Advanced", kAdvanced), XRAY_PAGE(L"About", kAbout),
        XRAY_PAGE(L"Notices", kNotices),
    };
#undef XRAY_PAGE
    constexpr int kPageCount = static_cast<int>(sizeof(kPages) / sizeof(kPages[0]));
    constexpr int kAboutPage = kPageCount - 2;     // About and Notices change no setting
    const glyph::Kind kPageGlyphs[kPageCount] = { glyph::Kind::Capture, glyph::Kind::Output,
                                                  glyph::Kind::Advanced, glyph::Kind::About,
                                                  glyph::Kind::Notices };

    const int kHeadingIds[] = { IDC_XLL_SEC1, IDC_VBA_SEC1, IDC_OUT_SEC1, IDC_ADV_SEC1, IDC_ADV_SEC2 };
    const int kRuleIds[]    = { IDC_XLL_RULE1, IDC_VBA_RULE1, IDC_OUT_RULE1, IDC_ADV_RULE1, IDC_ADV_RULE2 };
    const int kTitleIds[]   = { IDC_CAP_HDR, IDC_OUT_HDR, IDC_ADV_HDR, IDC_ABT_HDR, IDC_NOT_HDR };
    const int kCheckIds[]   = { IDC_XLL_ARGS, IDC_XLL_RET, IDC_VBA_ARGS, IDC_VBA_RET, IDC_VBA_OBJ };
    const int kButtonIds[]  = { IDOK, IDCANCEL, IDC_OUT_TAIL };
    const int kComboIds[]   = { IDC_XLL_DEPTH, IDC_VBA_DEPTH, IDC_OUT_FMT, IDC_ADV_FULL, IDC_ADV_LVL };
    const int kEditIds[]    = { IDC_OUT_DIR, IDC_OUT_FILE, IDC_ADV_BUF, IDC_ADV_LOG, IDC_ABT_LICENSE,
                                IDC_NOT_TEXT };

    template <size_t N> bool In(const int (&ids)[N], int id)
    {
        for (int i : ids) if (i == id) return true;
        return false;
    }

    // ---- the settings, as a draft: read on open, written on OK -------------------

    struct Draft
    {
        int  xllDepth = 0, vbaDepth = 0;
        bool xllArgs = false, xllRet = false;
        bool vbaArgs = false, vbaRet = false, vbaObj = false;
        bool pauseOnFull = true;
        int  format = 0;            // core::modes::Format
        wchar_t buffer[32] = {};
        int  logLevel = 0;
        bool armed = false;
    };

    const char* const kLevelNames[] = { "DEBUG", "INFO", "WARNING", "ERROR" };
    constexpr int kLevelCount = static_cast<int>(sizeof(kLevelNames) / sizeof(kLevelNames[0]));
    static_assert(static_cast<int>(core::Log::Level::Debug) == 0 &&
                  static_cast<int>(core::Log::Level::Error) == 3,
                  "kLevelNames no longer matches core::Log::Level");

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

    void ReadCurrent(Draft& d)
    {
        d.armed    = app::IsArmed();
        d.xllDepth = M::DepthIndex(L"ddXllDepth");
        d.vbaDepth = M::DepthIndex(L"ddVbaDepth");
        d.xllArgs  = M::ReadToggle(L"cbXllArgs");
        d.xllRet   = M::ReadToggle(L"cbXllRet");
        d.vbaArgs  = M::ReadToggle(L"cbVbaArgs");
        d.vbaRet   = M::ReadToggle(L"cbVbaRet");
        d.vbaObj   = M::ReadToggle(L"cbVbaObj");
        d.pauseOnFull = M::ReadToggle(L"cbPauseFull");
        d.format   = static_cast<int>(core::modes::GetFormat());
        M::BufferText(d.buffer, 32);
        const int level = static_cast<int>(core::Log::GetLevel());
        d.logLevel = (level >= 0 && level < kLevelCount) ? level : 1;
    }

    void Collect(HWND dlg, Draft& d)
    {
        d.xllDepth = static_cast<int>(SendDlgItemMessageW(dlg, IDC_XLL_DEPTH, CB_GETCURSEL, 0, 0));
        d.vbaDepth = static_cast<int>(SendDlgItemMessageW(dlg, IDC_VBA_DEPTH, CB_GETCURSEL, 0, 0));
        d.xllArgs  = IsChecked(dlg, IDC_XLL_ARGS);
        d.xllRet   = IsChecked(dlg, IDC_XLL_RET);
        d.vbaArgs  = IsChecked(dlg, IDC_VBA_ARGS);
        d.vbaRet   = IsChecked(dlg, IDC_VBA_RET);
        d.vbaObj   = IsChecked(dlg, IDC_VBA_OBJ);
        d.pauseOnFull = SendDlgItemMessageW(dlg, IDC_ADV_FULL, CB_GETCURSEL, 0, 0) == 0;
        d.format   = static_cast<int>(SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_GETCURSEL, 0, 0));
        GetDlgItemTextW(dlg, IDC_ADV_BUF, d.buffer, 32);
        d.logLevel = static_cast<int>(SendDlgItemMessageW(dlg, IDC_ADV_LVL, CB_GETCURSEL, 0, 0));
    }

    // False keeps the dialog open: the buffer size is the one field that can be wrong.
    bool Apply(const Draft& d, HWND dlg)
    {
        if (d.logLevel >= 0 && d.logLevel < kLevelCount)
        {
            core::Log::Level lvl;
            if (core::Log::LevelFromText(kLevelNames[d.logLevel], lvl)) core::Log::SetLevel(lvl);
        }
        if (app::IsArmed()) return true;    // the setters refuse everything else while armed

        M::SetDepthIndex(L"ddXllDepth", d.xllDepth);
        M::SetDepthIndex(L"ddVbaDepth", d.vbaDepth);
        M::WriteToggle(L"cbXllArgs",  d.xllArgs);
        M::WriteToggle(L"cbXllRet",   d.xllRet);
        M::WriteToggle(L"cbVbaArgs",  d.vbaArgs);
        M::WriteToggle(L"cbVbaRet",   d.vbaRet);
        M::WriteToggle(L"cbVbaObj",   d.vbaObj);
        M::WriteToggle(L"cbPauseFull", d.pauseOnFull);
        if (d.format == static_cast<int>(core::modes::Format::Csv) ||
            d.format == static_cast<int>(core::modes::Format::Jsonl))
            core::modes::SetFormat(static_cast<core::modes::Format>(d.format));
        if (M::SetBufferText(d.buffer)) return true;

        MessageBoxW(dlg,
            L"The buffer size was not understood.\r\n\r\n"
            L"Use a number with an optional unit: 64MB, 512KB, or 0 for "
            L"synchronous. The smallest ring is 16KB and the largest 240MB.",
            L"XRayXL Options", MB_OK | MB_ICONINFORMATION);
        return false;
    }

    // ---- palette -------------------------------------------------------------

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

    HBRUSH g_white = nullptr, g_pane = nullptr, g_hot = nullptr;
    HFONT  g_edit = nullptr;        // text boxes draw their own text: ClearType, as Excel's
    HFONT  g_editSoft = nullptr;    // the licence box: greyscale, as Excel's
    int    g_hotItem = -1;          // the category under the pointer
    int    g_page    = 0;           // the page showing
    bool   g_dropping = false;      // a drop-down is opening
    HWND   g_hotCtl  = nullptr;     // the control under the pointer

    // ---- scaling ---------------------------------------------------------------

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
    // Hairlines and outer margins step with the whole-number scale, as Excel's do.
    int Line(HWND h, int at96) { return at96 * (std::max)(1, DpiOf(h) / 96); }
    // Sizes Excel does not scale linearly: the line through its 100% and 150% values.
    int Fit(HWND h, int at96, int at144) { return at96 + MulDiv(at144 - at96, DpiOf(h) - 96, 48); }

    void Text(HWND h, HDC dc, const RECT& rc, const wchar_t* str, Face face, COLORREF colour, unsigned flags)
    {
        ui::text::Draw(dc, rc, str, face, colour, flags, DpiOf(h));
    }

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

    // ---- layout: the 96-DPI picture, in client pixels -------------------------------

    // w 0: to the content's right edge.  h 0: a drop-down.  h 21: a drop-down's label.
    struct Place { int id, x, y, w, h; };
    const Place kPlaces[] = {
        { IDC_PAGEICON,     182,  16,  28,  28 },
        { IDC_ARMEDNOTE,    169, 349,   0,  32 },

        { IDC_CAP_HDR,      222,  15,   0,  30 },
        { IDC_XLL_SEC1,     169,  64,   0,  20 }, { IDC_XLL_RULE1,  169,  87,   0,  1 },
        { IDC_XLL_DEPTHLBL, 182,  97, 118,  21 }, { IDC_XLL_DEPTH,  302,  97, 180,  0 },
        { IDC_XLL_ARGS,     182, 124,   0,  17 },
        { IDC_XLL_RET,      182, 146,   0,  17 },
        { IDC_VBA_SEC1,     169, 178,   0,  20 }, { IDC_VBA_RULE1,  169, 201,   0,  1 },
        { IDC_VBA_DEPTHLBL, 182, 211, 118,  21 }, { IDC_VBA_DEPTH,  302, 211, 180,  0 },
        { IDC_VBA_ARGS,     182, 238,   0,  17 },
        { IDC_VBA_RET,      182, 260,   0,  17 },
        { IDC_VBA_OBJ,      182, 282,   0,  17 },

        { IDC_OUT_HDR,      222,  15,   0,  30 },
        { IDC_OUT_SEC1,     169,  64,   0,  20 }, { IDC_OUT_RULE1,  169,  87,   0,  1 },
        { IDC_OUT_FMTLBL,   182,  97, 118,  21 }, { IDC_OUT_FMT,    302,  97, 180,  0 },
        { IDC_OUT_DIRLBL,   182, 133,   0,  15 }, { IDC_OUT_DIR,    182, 151,   0, 23 },
        { IDC_OUT_FILELBL,  182, 182,   0,  15 }, { IDC_OUT_FILE,   182, 200,   0, 23 },
        { IDC_OUT_TAIL,     182, 233, 130,  24 },

        { IDC_ADV_HDR,      222,  15,   0,  30 },
        { IDC_ADV_SEC1,     169,  64,   0,  20 }, { IDC_ADV_RULE1,  169,  87,   0,  1 },
        { IDC_ADV_BUFLBL,   182,  97, 118,  23 }, { IDC_ADV_BUF,    302,  97,  91, 23 },
        { IDC_ADV_BUFHINT,  302, 123,   0,  15 },
        { IDC_ADV_FULLLBL,  182, 146, 118,  21 }, { IDC_ADV_FULL,   302, 146, 180,  0 },
        { IDC_ADV_SEC2,     169, 182,   0,  20 }, { IDC_ADV_RULE2,  169, 205,   0,  1 },
        { IDC_ADV_LVLLBL,   182, 215, 118,  21 }, { IDC_ADV_LVL,    302, 215, 180,  0 },
        { IDC_ADV_LVLNOTE,  302, 242,   0,  15 },
        { IDC_ADV_LOGLBL,   182, 264,   0,  15 }, { IDC_ADV_LOG,    182, 282,   0, 23 },

        { IDC_ABT_HDR,      222,  15,   0,  30 },
        { IDC_NOT_HDR,      222,  15,   0,  30 },
    };
    constexpr int kClientW = 671, kClientH = 436;

    void LayoutFrame(HWND dlg)
    {
        auto place = [](HWND c, int x, int y, int w, int h) {
            SetWindowPos(c, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE); };

        // Sized from this DPI's frame metrics: mid-change the window still wears the old ones.
        RECT win, want{ 0, 0, Px(dlg, kClientW), Px(dlg, kClientH) };
        GetWindowRect(dlg, &win);
        typedef BOOL (WINAPI* AdjustFn)(LPRECT, DWORD, BOOL, DWORD, UINT);
        static const auto adjust = reinterpret_cast<AdjustFn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "AdjustWindowRectExForDpi")));
        const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_STYLE));
        const DWORD ex    = static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_EXSTYLE));
        if (adjust) adjust(&want, style, FALSE, ex, static_cast<UINT>(WindowDpi(dlg)));
        else        AdjustWindowRectEx(&want, style, FALSE, ex);
        const int winW = want.right - want.left, winH = want.bottom - want.top;
        place(dlg, (win.left + win.right - winW) / 2, (win.top + win.bottom - winH) / 2, winW, winH);
        RECT client;
        GetClientRect(dlg, &client);

        // Content scales from its left edge and from the pane's top; the margins do not.
        const int margin = Fit(dlg, 20, 26), paneW = Px(dlg, 107);
        const int left  = margin + paneW + margin;
        const int right = client.right - margin;
        const int top   = Line(dlg, 7);
        const int pitch = Fit(dlg, 22, 36);                  // check box to check box
        for (const Place& pl : kPlaces)
        {
            HWND c = GetDlgItem(dlg, pl.id);
            if (!c) continue;
            const int x = left + Px(dlg, pl.x - 169);
            int y = top + Px(dlg, pl.y - 7);
            if (In(kCheckIds, pl.id))
            {
                const int first = (pl.y >= 238) ? 238 : 124;         // each group's first box
                y = top + Px(dlg, first - 7) + (pl.y - first) / 22 * pitch;
            }
            int h = Px(dlg, pl.h);
            if (pl.h == 1)  h = Line(dlg, 1);
            if (pl.h == 0)  h = Px(dlg, 160);                // the open list's extent
            if (pl.h == 21) h = Fit(dlg, 21, 33);
            place(c, x, y, pl.w ? Px(dlg, pl.w) : right - x, h);
        }

        const int btnW = Px(dlg, 75), btnH = MulDiv(47, DpiOf(dlg), 192);    // 24, and 35 at 150%
        const int btnY = client.bottom - Line(dlg, 7) - btnH;
        const int cancelX = client.right - Line(dlg, 6) - btnW;
        place(GetDlgItem(dlg, IDCANCEL), cancelX, btnY, btnW, btnH);
        place(GetDlgItem(dlg, IDOK), cancelX - Px(dlg, 9) - btnW, btnY, btnW, btnH);
        const int paneBottom = btnY - margin;
        place(GetDlgItem(dlg, IDC_CATEGORIES), margin, top, paneW, paneBottom - top);

        // About: half-line gaps, and the licence box runs down to the pane's bottom edge.
        const int x = left + Px(dlg, 182 - 169), w = right - x, lineH = Px(dlg, 15), gap = Px(dlg, 8);
        const int firstY = top + Px(dlg, 61 - 7);
        int y = firstY;
        place(GetDlgItem(dlg, IDC_ABT_OWNER), x, y, w, lineH);
        y += lineH + gap + lineH;
        place(GetDlgItem(dlg, IDC_ABT_LICLBL), x, y, w, lineH);
        y += lineH + Px(dlg, 3);
        place(GetDlgItem(dlg, IDC_ABT_LICENSE), x, y, w, paneBottom - y);

        // Notices: the one box, from the first line down to the same edge.
        place(GetDlgItem(dlg, IDC_NOT_TEXT), x, firstY, w, paneBottom - firstY);
    }

    HFONT EditFont(HWND dlg, DWORD quality)
    {
        return CreateFontW(-MulDiv(9, DpiOf(dlg), 72), 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
                           DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, quality,
                           DEFAULT_PITCH | FF_DONTCARE, L"Segoe UI");
    }

    // Everything that depends on the DPI: at open, and when dragged to another monitor.
    void ApplyDpi(HWND dlg)
    {
        LayoutFrame(dlg);

        if (g_edit)     DeleteObject(g_edit);
        if (g_editSoft) DeleteObject(g_editSoft);
        g_edit     = EditFont(dlg, DEFAULT_QUALITY);
        g_editSoft = EditFont(dlg, ANTIALIASED_QUALITY);
        for (int id : kEditIds)
        {
            HWND e = GetDlgItem(dlg, id);
            const HFONT f = (id == IDC_ABT_LICENSE || id == IDC_NOT_TEXT) ? g_editSoft : g_edit;
            SendMessageW(e, WM_SETFONT, reinterpret_cast<WPARAM>(f), FALSE);
            SendMessageW(e, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, 0);
            SetWindowPos(e, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
                                                 SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }
        for (int id : kComboIds)
        {
            // a closed drop-down is its item height plus 6
            SendDlgItemMessageW(dlg, id, CB_SETITEMHEIGHT, static_cast<WPARAM>(-1), Fit(dlg, 21, 33) - 6);
            SendDlgItemMessageW(dlg, id, CB_SETITEMHEIGHT, 0, Px(dlg, 24));
        }
        SendDlgItemMessageW(dlg, IDC_CATEGORIES, LB_SETITEMHEIGHT, 0, Px(dlg, 27));

        RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
    }

    // ---- drawing ---------------------------------------------------------------

    void DrawLabel(const DRAWITEMSTRUCT* di)
    {
        const int id = static_cast<int>(di->CtlID);
        wchar_t text[256] = {};
        GetWindowTextW(di->hwndItem, text, 256);
        const bool heading = In(kHeadingIds, id), title = In(kTitleIds, id);
        const bool oneLine = (di->rcItem.bottom - di->rcItem.top) < Px(di->hwndItem, 28);
        unsigned flags = heading ? 0u : (title || oneLine) ? ui::text::kVCentre : ui::text::kWrap;
        const Face face = heading ? Face::Bold : title ? Face::Title : Face::Body;
        Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
            FillRect(dc, &rc, g_white);
            Text(di->hwndItem, dc, rc, text, face, kText, flags); });
    }

    void DrawPageIcon(const DRAWITEMSTRUCT* di)
    {
        HWND h = di->hwndItem;
        Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& r) {
            FillRect(dc, &r, g_white);
            glyph::Draw(dc, r, kPageGlyphs[g_page], Line(h, 1), kIconInk, kPage); });
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

    // The pane's frame curves into the client corners, so rows and the space below redraw it.
    void DrawPaneFrame(HWND h, HDC dc, const RECT& windowRc)
    {
        ui::soft::RoundRect(dc, windowRc, ui::soft::Box{ 0, kEdge, Line(h, 1), Px(h, 4), false });
    }

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

    // ---- subclasses: hover, and the frames the stock controls cannot draw ------------

    const wchar_t* const kOldProc = L"XRayOldProc";

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

    void InvalidateItem(HWND list, int item)
    {
        RECT r;
        if (item >= 0 && SendMessageW(list, LB_GETITEMRECT, item, reinterpret_cast<LPARAM>(&r)) != LB_ERR)
            InvalidateRect(list, &r, FALSE);
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
            // The space under the rows, and the frame's curve through it.
            HRGN upd = CreateRectRgn(0, 0, 0, 0);
            const int kind = GetUpdateRgn(h, upd, FALSE);
            const LRESULT res = Previous(h, m, wp, lp);
            RECT c, last{};
            GetClientRect(h, &c);
            const LRESULT n = SendMessageW(h, LB_GETCOUNT, 0, 0);
            if (n > 0) SendMessageW(h, LB_GETITEMRECT, n - 1, reinterpret_cast<LPARAM>(&last));
            if (kind != ERROR && kind != NULLREGION && last.bottom < c.bottom)
            {
                HDC dc = GetDC(h);
                SelectClipRgn(dc, upd);
                RECT below = c;
                below.top = last.bottom;
                RECT frame = c;
                InflateRect(&frame, Line(h, 1), Line(h, 1));
                OffsetRect(&frame, 0, -below.top);
                Buffered(dc, below, [&](HDC mem, const RECT& rc) {
                    FillRect(mem, &rc, g_pane);
                    DrawPaneFrame(h, mem, frame); });
                ReleaseDC(h, dc);
            }
            DeleteObject(upd);
            return res;
        }
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

    // Text boxes: the stock edit edits, and Excel's frame is painted in a margin around it.
    bool IsMultiline(HWND h) { return (GetWindowLongPtrW(h, GWL_STYLE) & ES_MULTILINE) != 0; }
    bool EditLit(HWND h)
    {
        return IsWindowEnabled(h) && !IsMultiline(h) && (g_hotCtl == h || GetFocus() == h);
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

    // The two path boxes offer what is useful for a path, not the edit control's own menu.
    bool PathMenu(HWND edit, LPARAM lp)
    {
        const int id = GetDlgCtrlID(edit);
        if (id != IDC_OUT_DIR && id != IDC_OUT_FILE && id != IDC_ADV_LOG) return false;
        const bool isFile = (id != IDC_OUT_DIR);
        const std::wstring path = (id == IDC_OUT_FILE) ? emit::csv::Path()
                                : (id == IDC_ADV_LOG)  ? core::Log::Path() : core::EnsureAppSubdir(L"TraceFiles");
        const bool exists = !path.empty() && GetFileAttributesW(path.c_str()) != INVALID_FILE_ATTRIBUTES;

        enum { Copy = 1, Open, Reveal, Tail };
        HMENU menu = CreatePopupMenu();
        const auto add = [&](UINT cmd, const wchar_t* text, bool on) {
            AppendMenuW(menu, MF_STRING | (on ? 0u : MF_GRAYED), cmd, text); };
        add(Copy, L"Copy Path", !path.empty());
        if (isFile)
        {
            add(Reveal, L"Reveal in File Explorer", exists);
            add(Tail, L"Tail", !path.empty());
        }
        else add(Open, L"Open in File Explorer", true);

        POINT at{ static_cast<short>(LOWORD(lp)), static_cast<short>(HIWORD(lp)) };
        if (at.x == -1 && at.y == -1)               // from the keyboard
        {
            RECT r;
            GetWindowRect(edit, &r);
            at = POINT{ r.left, r.bottom };
        }
        HWND dlg = GetParent(edit);
        const int cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, at.x, at.y, 0, dlg, nullptr);
        DestroyMenu(menu);

        trace::Result r = trace::Result::Ok;
        if (cmd == Copy)   r = trace::CopyText(dlg, path);
        if (cmd == Open)   r = trace::OpenFolder(path);
        if (cmd == Reveal) r = trace::RevealFile(path);
        if (cmd == Tail)   r = trace::TailInPowerShell(path);
        if (r != trace::Result::Ok) MessageBoxW(dlg, trace::Explain(r), L"XRayXL", MB_OK | MB_ICONINFORMATION);
        return true;
    }

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
            if (PathMenu(h, lp)) return 0;
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

    // ---- the system's parts: caption and the drop-down's popup ---------------------

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

    // ---- page contents -----------------------------------------------------------

    HMODULE Self()
    {
        HMODULE self = nullptr;
        GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
                           reinterpret_cast<LPCWSTR>(&Self), &self);
        return self;
    }

    // A compiled-in text file, reflowed so the box can wrap it.
    std::wstring ResourceText(int id)
    {
        HRSRC res = FindResourceW(Self(), MAKEINTRESOURCEW(id), RT_RCDATA);
        HGLOBAL mem = res ? LoadResource(Self(), res) : nullptr;
        const char* bytes = mem ? static_cast<const char*>(LockResource(mem)) : nullptr;
        if (!bytes) return L"";
        const std::string out = ui::text::Reflow(std::string(bytes, SizeofResource(Self(), res)));

        std::wstring wide(out.size(), L'\0');
        const int got = MultiByteToWideChar(CP_UTF8, 0, out.data(), static_cast<int>(out.size()),
                                            &wide[0], static_cast<int>(wide.size()));
        wide.resize(got > 0 ? got : 0);
        return wide;
    }

    void FillPages(HWND dlg)
    {
        SetDlgItemTextW(dlg, IDC_OUT_DIR, core::EnsureAppSubdir(L"TraceFiles").c_str());
        const std::wstring file = emit::csv::Path();
        SetDlgItemTextW(dlg, IDC_OUT_FILE, file.empty() ? L"(named when tracing is armed)" : file.c_str());

        wchar_t line[MAX_PATH * 2 + 32];
        _snwprintf_s(line, _TRUNCATE, L"About XRayXL v%hs (%dBit)", app::VersionText(), static_cast<int>(sizeof(void*) * 8));
        SetDlgItemTextW(dlg, IDC_ABT_HDR, line);
        SetDlgItemTextW(dlg, IDC_ABT_OWNER, L"\u00A9 2026 Andrew Lockhart");
        SetDlgItemTextW(dlg, IDC_ADV_LOG, core::Log::Path().c_str());
        SetDlgItemTextW(dlg, IDC_ABT_LICENSE, ResourceText(IDR_LICENSE).c_str());
        SetDlgItemTextW(dlg, IDC_NOT_TEXT, ResourceText(IDR_NOTICES).c_str());
    }

    void ShowPage(HWND dlg, int page)
    {
        for (int p = 0; p < kPageCount; ++p)
            for (int i = 0; i < kPages[p].count; ++i)
                ShowWindow(GetDlgItem(dlg, kPages[p].ids[i]), (p == page) ? SW_SHOW : SW_HIDE);
        ShowWindow(GetDlgItem(dlg, IDC_ARMEDNOTE), page >= kAboutPage ? SW_HIDE : SW_SHOW);
        g_page = page;
        InvalidateRect(GetDlgItem(dlg, IDC_PAGEICON), nullptr, FALSE);
    }

    // Greyed by the rule the setters refuse by.
    void ApplyArmedState(HWND dlg, bool armed)
    {
        const int locked[] = { IDC_XLL_DEPTH, IDC_XLL_ARGS, IDC_XLL_RET,
                               IDC_VBA_DEPTH, IDC_VBA_ARGS, IDC_VBA_RET, IDC_VBA_OBJ,
                               IDC_ADV_BUF, IDC_ADV_FULL, IDC_OUT_FMT };
        for (int id : locked) EnableWindow(GetDlgItem(dlg, id), armed ? FALSE : TRUE);
        SetDlgItemTextW(dlg, IDC_ARMEDNOTE, armed
            ? L"Tracing is armed. These settings are read when a session starts, so "
              L"disarm first to change them. The log level can still be changed."
            : L"");
    }

    void InitControls(HWND dlg, const Draft& d)
    {
        for (const Page& p : kPages)
            SendDlgItemMessageW(dlg, IDC_CATEGORIES, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(p.name));
        SendDlgItemMessageW(dlg, IDC_CATEGORIES, LB_SETCURSEL, 0, 0);

        for (const wchar_t* t : { L"Off", L"Top level only", L"All calls" })
        {
            SendDlgItemMessageW(dlg, IDC_XLL_DEPTH, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(t));
            SendDlgItemMessageW(dlg, IDC_VBA_DEPTH, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(t));
        }
        for (const wchar_t* t : { L"Pause (lose nothing)", L"Drop (never wait)" })
            SendDlgItemMessageW(dlg, IDC_ADV_FULL, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(t));
        // In core::modes::Format order.
        for (const wchar_t* t : { L"CSV", L"JSON Lines" })
            SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(t));
        for (const char* t : kLevelNames)
        {
            wchar_t w[16];
            core::WidenUtf8(t, w, 16);
            SendDlgItemMessageW(dlg, IDC_ADV_LVL, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(w));
        }

        SendDlgItemMessageW(dlg, IDC_XLL_DEPTH, CB_SETCURSEL, d.xllDepth, 0);
        SendDlgItemMessageW(dlg, IDC_VBA_DEPTH, CB_SETCURSEL, d.vbaDepth, 0);
        SetChecked(dlg, IDC_XLL_ARGS, d.xllArgs);
        SetChecked(dlg, IDC_XLL_RET,  d.xllRet);
        SetChecked(dlg, IDC_VBA_ARGS, d.vbaArgs);
        SetChecked(dlg, IDC_VBA_RET,  d.vbaRet);
        SetChecked(dlg, IDC_VBA_OBJ,  d.vbaObj);
        SendDlgItemMessageW(dlg, IDC_ADV_FULL, CB_SETCURSEL, d.pauseOnFull ? 0 : 1, 0);
        SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_SETCURSEL, d.format, 0);
        SetDlgItemTextW(dlg, IDC_ADV_BUF, d.buffer);
        SendDlgItemMessageW(dlg, IDC_ADV_LVL, CB_SETCURSEL, d.logLevel, 0);

        // every text static is owner-drawn, so its text goes through Text()
        for (HWND c = GetWindow(dlg, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT))
        {
            wchar_t cls[16] = {};
            GetClassNameW(c, cls, 16);
            const LONG_PTR st = GetWindowLongPtrW(c, GWL_STYLE);
            if (!lstrcmpiW(cls, L"Static") && (st & SS_TYPEMASK) == SS_LEFT)
                SetWindowLongPtrW(c, GWL_STYLE, (st & ~static_cast<LONG_PTR>(SS_TYPEMASK)) | SS_OWNERDRAW);
        }

        Subclass(GetDlgItem(dlg, IDC_CATEGORIES), ListProc);
        SetWindowPos(GetDlgItem(dlg, IDC_CATEGORIES), nullptr, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        for (int id : kButtonIds) Subclass(GetDlgItem(dlg, id), ButtonProc);
        for (int id : kCheckIds)  Subclass(GetDlgItem(dlg, id), ButtonProc);
        for (int id : kEditIds)   Subclass(GetDlgItem(dlg, id), EditProc);
        for (int id : kComboIds)
        {
            Subclass(GetDlgItem(dlg, id), ComboProc);
            StyleDropList(GetDlgItem(dlg, id));
        }
    }

    INT_PTR CALLBACK Proc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
    {
        static Draft s_draft;

        switch (msg)
        {
        case WM_INITDIALOG:
        {
            g_hotItem = -1;
            g_hotCtl  = nullptr;
            g_dropping = false;
            g_white = CreateSolidBrush(kPage);
            g_pane  = CreateSolidBrush(kPane);
            g_hot   = CreateSolidBrush(kHotFace);

            ReadCurrent(s_draft);
            InitControls(dlg, s_draft);
            FillPages(dlg);
            WhiteCaption(dlg);
            SendMessageW(dlg, DM_SETDEFID, IDOK, 0);    // an owner-drawn button cannot say "default" itself

            // The layout is ours in pixels, so the dialog manager's own DPI rescale is switched off.
            typedef BOOL (WINAPI* BehaviourFn)(HWND, int, int);
            const auto behaviour = reinterpret_cast<BehaviourFn>(reinterpret_cast<void*>(
                GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetDialogDpiChangeBehavior")));
            if (behaviour) behaviour(dlg, 7 /*DDC_DISABLE_ALL*/, 7);
            ApplyDpi(dlg);

            ApplyArmedState(dlg, s_draft.armed);
            ShowPage(dlg, 0);
            return TRUE;
        }

        case WM_DPICHANGED:
        {
            const RECT* to = reinterpret_cast<const RECT*>(lp);
            if (to)
                SetWindowPos(dlg, nullptr, to->left, to->top, to->right - to->left, to->bottom - to->top,
                             SWP_NOZORDER | SWP_NOACTIVATE);
            ApplyDpi(dlg);
            return TRUE;
        }

        case WM_MEASUREITEM:
        {
            MEASUREITEMSTRUCT* mi = reinterpret_cast<MEASUREITEMSTRUCT*>(lp);
            if (mi) mi->itemHeight = static_cast<UINT>(Px(dlg, mi->CtlID == IDC_CATEGORIES ? 27 : 24));
            return TRUE;
        }

        case WM_DRAWITEM:
        {
            const DRAWITEMSTRUCT* di = reinterpret_cast<const DRAWITEMSTRUCT*>(lp);
            if (!di) return FALSE;
            const int id = static_cast<int>(di->CtlID);
            if (id == IDC_CATEGORIES)         DrawCategory(di);
            else if (id == IDC_PAGEICON)      DrawPageIcon(di);
            else if (In(kCheckIds, id))       DrawCheck(di);
            else if (In(kRuleIds, id))        DrawRule(di);
            else if (In(kButtonIds, id))      DrawButton(di);
            else if (In(kComboIds, id))       DrawComboItem(di);
            else if (di->CtlType == ODT_STATIC) DrawLabel(di);
            else return FALSE;
            return TRUE;
        }

        case WM_CTLCOLORDLG:
            return reinterpret_cast<INT_PTR>(g_white);
        case WM_CTLCOLORLISTBOX:
        {
            const bool categories = reinterpret_cast<HWND>(lp) == GetDlgItem(dlg, IDC_CATEGORIES);
            SetBkColor(reinterpret_cast<HDC>(wp), categories ? kPane : kPage);
            return reinterpret_cast<INT_PTR>(categories ? g_pane : g_white);
        }
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        {
            // an edit paints its own background, so it needs its state's colour
            HDC dc = reinterpret_cast<HDC>(wp);
            const bool lit = EditLit(reinterpret_cast<HWND>(lp));
            SetTextColor(dc, kText);
            SetBkColor(dc, lit ? kHotFace : kPage);
            return reinterpret_cast<INT_PTR>(lit ? g_hot : g_white);
        }

        case WM_COMMAND:
        {
            const int id = LOWORD(wp);
            if (id == IDC_CATEGORIES && HIWORD(wp) == LBN_SELCHANGE)
            {
                const int sel = static_cast<int>(SendDlgItemMessageW(dlg, IDC_CATEGORIES, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < kPageCount) ShowPage(dlg, sel);
                return TRUE;
            }
            if (In(kComboIds, id) && (HIWORD(wp) == CBN_DROPDOWN || HIWORD(wp) == CBN_CLOSEUP))
            {
                g_dropping = HIWORD(wp) == CBN_DROPDOWN;
                return TRUE;
            }
            if (HIWORD(wp) == BN_CLICKED && In(kCheckIds, id))
            {
                SetChecked(dlg, id, !IsChecked(dlg, id));   // an owner-drawn box has no state of its own
                return TRUE;
            }
            if (id == IDC_OUT_TAIL)
            {
                const trace::Result r = trace::TailInPowerShell(emit::csv::Path());
                if (r != trace::Result::Ok)
                    MessageBoxW(dlg, trace::Explain(r), L"XRayXL", MB_OK | MB_ICONINFORMATION);
                return TRUE;
            }
            if (id == IDOK)
            {
                Draft d = s_draft;
                Collect(dlg, d);
                if (Apply(d, dlg)) EndDialog(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL) { EndDialog(dlg, IDCANCEL); return TRUE; }
            return FALSE;
        }

        case WM_CLOSE:
            EndDialog(dlg, IDCANCEL);
            return TRUE;

        case WM_DESTROY:
            for (HBRUSH* b : { &g_white, &g_pane, &g_hot }) if (*b) { DeleteObject(*b); *b = nullptr; }
            for (HFONT* f : { &g_edit, &g_editSoft })       if (*f) { DeleteObject(*f); *f = nullptr; }
            ui::text::Release();
            return FALSE;
        }
        return FALSE;
    }

    // For XRayXL_Options, which has no ribbon control to ask for a window.
    HWND ExcelMainWindow()
    {
        const DWORD mine = GetCurrentProcessId();
        HWND firstOfOurs = nullptr;
        for (HWND h = nullptr; (h = FindWindowExW(nullptr, h, L"XLMAIN", nullptr)) != nullptr; )
        {
            DWORD owner = 0;
            GetWindowThreadProcessId(h, &owner);
            if (owner != mine) continue;
            if (IsWindowVisible(h)) return h;
            if (!firstOfOurs) firstOfOurs = h;
        }
        return firstOfOurs;
    }
}

bool Show(void* ownerHwnd)
{
    HWND owner = static_cast<HWND>(ownerHwnd);
    if (!owner) owner = ExcelMainWindow();

    if (!ui::text::Ready())
    {
        core::Log::Warning("options: Direct2D/DirectWrite unavailable; the dialog cannot draw its text");
        return false;
    }

    // Office calls add-ins system-DPI-aware, which blurs a window on a scaled monitor.
    typedef HANDLE (WINAPI* SetContextFn)(HANDLE);
    const auto setContext = reinterpret_cast<SetContextFn>(reinterpret_cast<void*>(
        GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetThreadDpiAwarenessContext")));
    const HANDLE perMonitorV2 = reinterpret_cast<HANDLE>(static_cast<INT_PTR>(-4));
    const HANDLE before = setContext ? setContext(perMonitorV2) : nullptr;

    const INT_PTR r = DialogBoxParamW(Self(), MAKEINTRESOURCEW(IDD_XRAY_OPTIONS), owner, Proc, 0);
    if (before) setContext(before);

    if (r == -1)
    {
        char line[64];
        _snprintf_s(line, _TRUNCATE, "options: dialog failed to open (%lu)", GetLastError());
        core::Log::Warning(line);
        return false;
    }
    core::Log::Note(r == IDOK ? "options: applied" : "options: cancelled");
    return r == IDOK;
}
}
}
