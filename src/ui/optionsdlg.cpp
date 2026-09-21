#include "optionsdlg.h"
#include "optionsres.h"
#include "dlgexcelstyle.h"
#include "ribbonmodel.h"
#include "traceactions.h"
#include "glyphs.h"
#include "reflow.h"

#include <algorithm>

#include "app/session.h"
#include "app/settings.h"
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
    using namespace ui::excelstyle;
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

    // ---- the settings, as a draft: read on open, written on Apply ----------------

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

    // Whether Apply would change anything. A buffer size that does not parse counts as a change,
    // so Apply is what says why it is wrong.
    bool Differs(const Draft& d, const Draft& cur)
    {
        if (d.logLevel != cur.logLevel) return true;
        if (cur.armed) return false;        // the setters refuse everything else while armed
        std::size_t bytes = 0;
        if (!M::ParseBufferBox(d.buffer, bytes) || bytes != core::modes::GetBufferBytes()) return true;
        return d.xllDepth != cur.xllDepth || d.vbaDepth != cur.vbaDepth ||
               d.xllArgs != cur.xllArgs || d.xllRet != cur.xllRet ||
               d.vbaArgs != cur.vbaArgs || d.vbaRet != cur.vbaRet || d.vbaObj != cur.vbaObj ||
               d.pauseOnFull != cur.pauseOnFull || d.format != cur.format;
    }

    // The dialog's values over the live settings, so the comparison is always with what is current.
    void Pending(HWND dlg, Draft& d, Draft& cur)
    {
        ReadCurrent(cur);
        d = cur;
        Collect(dlg, d);
    }

    void UpdateApply(HWND dlg)
    {
        Draft d, cur;
        Pending(dlg, d, cur);
        EnableWindow(GetDlgItem(dlg, IDOK), Differs(d, cur) ? TRUE : FALSE);
    }

    // False keeps the dialog open, with nothing applied: the buffer size is the one field that can
    // be wrong, so it is checked first.
    bool Apply(const Draft& d, HWND dlg)
    {
        std::size_t bytes = 0;
        if (!d.armed && !M::ParseBufferBox(d.buffer, bytes))
        {
            MessageBoxW(dlg,
                L"The buffer size was not understood.\r\n\r\n"
                L"Use a number with an optional unit: 64MB, 512KB, or 0 for "
                L"synchronous. The smallest ring is 16KB and the largest 240MB.",
                L"XRayXL Options", MB_OK | MB_ICONINFORMATION);
            return false;
        }
        if (d.logLevel >= 0 && d.logLevel < kLevelCount)
        {
            core::Log::Level lvl;
            if (core::Log::LevelFromText(kLevelNames[d.logLevel], lvl)) core::Log::SetLevel(lvl);
        }
        if (d.armed) return true;           // the setters refuse everything else while armed

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
        core::modes::SetBufferBytes(bytes);
        return true;
    }

    int g_page = 0;                 // the page showing, for the glyph

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

        // One margin all round, as the Diagnostics dialog has: Cancel's right edge lines up
        // with the content's, and the gap beneath it matches the gap above.
        const int btnW = Px(dlg, 75), btnH = MulDiv(47, DpiOf(dlg), 192);    // 24, and 35 at 150%
        const int btnY = client.bottom - margin - btnH;
        const int cancelX = right - btnW;
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

    // Everything that depends on the DPI: at open, and when dragged to another monitor.
    void ApplyDpi(HWND dlg)
    {
        LayoutFrame(dlg);

        MakeFonts(dlg);
        for (int id : kEditIds)
        {
            HWND e = GetDlgItem(dlg, id);
            const HFONT f = (id == IDC_ABT_LICENSE || id == IDC_NOT_TEXT) ? SoftFont() : FieldFont();
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

    void DrawPageIcon(const DRAWITEMSTRUCT* di)
    {
        HWND h = di->hwndItem;
        Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& r) {
            FillRect(dc, &r, PageBrush());
            glyph::Draw(dc, r, kPageGlyphs[g_page], Line(h, 1), kIconInk, kPage); });
    }

    // ---- subclasses: hover, and the frames the stock controls cannot draw ------------

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

    // ---- page contents -----------------------------------------------------------

    // A compiled-in text file, reflowed so the box can wrap it.
    std::wstring ResourceText(int id)
    {
        HRSRC res = FindResourceW(excelstyle::Self(), MAKEINTRESOURCEW(id), RT_RCDATA);
        HGLOBAL mem = res ? LoadResource(excelstyle::Self(), res) : nullptr;
        const char* bytes = mem ? static_cast<const char*>(LockResource(mem)) : nullptr;
        if (!bytes) return L"";
        const std::string out = ui::text::Reflow(std::string(bytes, SizeofResource(excelstyle::Self(), res)));

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
            excelstyle::Begin();
            excelstyle::SetEditMenu(PathMenu);

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
            UpdateApply(dlg);
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
            else if (di->CtlType == ODT_STATIC)
            {
                const bool heading = In(kHeadingIds, id), title = In(kTitleIds, id);
                const bool oneLine = (di->rcItem.bottom - di->rcItem.top) < Px(di->hwndItem, 28);
                DrawLabel(di, heading ? Face::Bold : title ? Face::Title : Face::Body,
                          heading ? 0u : (title || oneLine) ? ui::text::kVCentre : ui::text::kWrap);
            }
            else return FALSE;
            return TRUE;
        }

        case WM_CTLCOLORDLG:
            return reinterpret_cast<INT_PTR>(PageBrush());
        case WM_CTLCOLORLISTBOX:
        {
            const bool categories = reinterpret_cast<HWND>(lp) == GetDlgItem(dlg, IDC_CATEGORIES);
            SetBkColor(reinterpret_cast<HDC>(wp), categories ? kPane : kPage);
            return reinterpret_cast<INT_PTR>(categories ? PaneBrush() : PageBrush());
        }
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        {
            // an edit paints its own background, so it needs its state's colour
            HDC dc = reinterpret_cast<HDC>(wp);
            const bool lit = EditLit(reinterpret_cast<HWND>(lp));
            SetTextColor(dc, kText);
            SetBkColor(dc, lit ? kHotFace : kPage);
            return reinterpret_cast<INT_PTR>(lit ? HotBrush() : PageBrush());
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
                excelstyle::SetDropping(HIWORD(wp) == CBN_DROPDOWN);
                return TRUE;
            }
            if (In(kComboIds, id) && HIWORD(wp) == CBN_SELCHANGE) { UpdateApply(dlg); return TRUE; }
            if (id == IDC_ADV_BUF && HIWORD(wp) == EN_CHANGE)    { UpdateApply(dlg); return TRUE; }
            if (HIWORD(wp) == BN_CLICKED && In(kCheckIds, id))
            {
                SetChecked(dlg, id, !IsChecked(dlg, id));   // an owner-drawn box has no state of its own
                UpdateApply(dlg);
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
                // Enter reaches here with Apply greyed, so the check is made again.
                Draft d, cur;
                Pending(dlg, d, cur);
                if (!Differs(d, cur)) return TRUE;
                const app::settings::Snapshot before = app::settings::Take();
                if (!Apply(d, dlg)) return TRUE;
                const std::string changed = app::settings::Changes(before, app::settings::Take());
                core::Log::Note("options: applied -- " + (changed.empty() ? std::string("nothing changed") : changed));
                EndDialog(dlg, IDOK);
                return TRUE;
            }
            if (id == IDCANCEL) { EndDialog(dlg, IDCANCEL); return TRUE; }
            return FALSE;
        }

        case WM_CLOSE:
            EndDialog(dlg, IDCANCEL);
            return TRUE;

        case WM_DESTROY:
            excelstyle::End();
            return FALSE;
        }
        return FALSE;
    }

}

bool Show(void* ownerHwnd)
{
    // The ribbon is the only caller and it hands over its own window; a null owner leaves the
    // dialog unowned, which is what the tests open it as.
    HWND owner = static_cast<HWND>(ownerHwnd);

    if (!ui::text::Ready())
    {
        core::Log::Warning("options: Direct2D/DirectWrite unavailable; the dialog cannot draw its text");
        return false;
    }

    core::Log::Note(std::string("options: opened") + (app::IsArmed() ? " while armed" : "") +
                    " -- " + app::settings::List(app::settings::Take()));
    const excelstyle::DpiScope perMonitor;
    const INT_PTR r = DialogBoxParamW(excelstyle::Self(), MAKEINTRESOURCEW(IDD_XRAY_OPTIONS), owner, Proc, 0);

    if (r == -1)
    {
        char line[64];
        _snprintf_s(line, _TRUNCATE, "options: dialog failed to open (%lu)", GetLastError());
        core::Log::Warning(line);
        return false;
    }
    if (r != IDOK) core::Log::Note("options: cancelled");
    return r == IDOK;
}
}
}
