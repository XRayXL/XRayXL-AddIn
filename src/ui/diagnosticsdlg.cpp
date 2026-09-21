#include "diagnosticsdlg.h"
#include "diagres.h"
#include "dlgexcelstyle.h"
#include "gridlist.h"
#include "glyphs.h"
#include "traceactions.h"

#include "diag/snapshot.h"
#include "diag/tableout.h"

#include "app/session.h"
#include "core/log.h"
#include "emit/csv.h"

#include <windows.h>
#include <windowsx.h>
#include <algorithm>
#include <cwchar>
#include <string>

// The frame and controls are the Options dialog's; the list is Excel's Name Manager. Unlike
// Options this one resizes, so the layout reads the client rect rather than setting it.

namespace ui
{
namespace diagnostics
{
namespace
{
    using namespace ui::excelstyle;
    using ui::text::Face;

    struct PageDef
    {
        const wchar_t* name;
        const wchar_t* heading;
        glyph::Kind    glyph;
        int            pathColumn;     // a column Explorer can reveal, or -1
    };

    const PageDef kPages[] = {
        { L"Modules",     L"Every module loaded into this Excel process.",
          glyph::Kind::Modules,     4 },
        { L"Environment", L"The environment this Excel process was started with.",
          glyph::Kind::Environment, -1 },
        { L"Process",     L"What this Excel process is using, and what it is running on.",
          glyph::Kind::Process,     -1 },
    };
    constexpr int kPageCount = static_cast<int>(sizeof(kPages) / sizeof(kPages[0]));

    const int kTitleIds[] = { IDC_DIAG_HDR };
    const int kEditIds[]  = { IDC_DIAG_SEARCH };

    template <size_t N> bool In(const int (&ids)[N], int id)
    {
        for (int i : ids) if (i == id) return true;
        return false;
    }

    int  g_page = 0;
    HWND g_list = nullptr;
    RECT g_grip{};        // where the grip was last drawn, so a resize can erase it

    // All three pages, read once when the dialog opens. The dialog is modal and shows nothing
    // live, so one reading of the process is what every page should describe.
    diag::Table g_snapshot[kPageCount];

    void ReadEverything()
    {
        diag::Session s;
        s.version   = app::VersionText();
        s.armed     = app::IsArmed();
        s.traceFile = emit::csv::Path();
        s.logFile   = core::Log::Path();

        g_snapshot[0] = diag::Modules();
        g_snapshot[1] = diag::Environment();
        g_snapshot[2] = diag::Process(s);
    }

    void LoadPage(HWND dlg, int page)
    {
        g_page = page;
        SetDlgItemTextW(dlg, IDC_DIAG_HDR, kPages[page].heading);
        SetDlgItemTextW(dlg, IDC_DIAG_SEARCH, L"");
        grid::SetFilter(g_list, L"");
        grid::SetTable(g_list, g_snapshot[page]);
        InvalidateRect(GetDlgItem(dlg, IDC_DIAG_PAGEICON), nullptr, FALSE);
    }

    // ---- layout ----------------------------------------------------------------------------
    // One margin all round, and nothing under the list: it runs to the bottom margin, as the
    // category pane runs to the same line.

    constexpr int kClientW = 900, kClientH = 433;   // the size it opens at
    constexpr int kMinW    = 620, kMinH    = 260;

    void Layout(HWND dlg)
    {
        HDWP defer = BeginDeferWindowPos(8);
        auto place = [&defer](HWND c, int x, int y, int w, int h) {
            if (c && defer)
                defer = DeferWindowPos(defer, c, nullptr, x, y, w, h, SWP_NOZORDER | SWP_NOACTIVATE); };

        RECT client;
        GetClientRect(dlg, &client);

        const int margin = Fit(dlg, 20, 26), paneW = Px(dlg, 107);
        const int left  = margin + paneW + margin;
        const int right = client.right - margin;
        const int top   = Line(dlg, 7);
        const int fieldH = Fit(dlg, 23, 35);
        const int bottom = client.bottom - margin;

        place(GetDlgItem(dlg, IDC_DIAG_CATEGORIES), margin, top, paneW, bottom - top);

        // the content starts one margin from the pane, as it does from every other edge
        const int x = left, w = right - x;
        const int iconW = Px(dlg, 28);
        place(GetDlgItem(dlg, IDC_DIAG_PAGEICON), x, top + Px(dlg, 9), iconW, iconW);
        place(GetDlgItem(dlg, IDC_DIAG_HDR), x + iconW + Px(dlg, 12), top + Px(dlg, 8),
              w - iconW - Px(dlg, 12), Px(dlg, 30));

        const int searchY = top + Px(dlg, 50);
        const int labelW = Px(dlg, 56);
        place(GetDlgItem(dlg, IDC_DIAG_SEARCHLBL), x, searchY, labelW, fieldH);
        place(GetDlgItem(dlg, IDC_DIAG_SEARCH), x + labelW, searchY, w - labelW, fieldH);

        const int listY = searchY + fieldH + Px(dlg, 10);
        place(g_list, x, listY, w, (std::max)(Px(dlg, 60), bottom - listY));

        if (defer) EndDeferWindowPos(defer);

        // The grip is anchored to the corner, so a resize moves it and the old one has to go.
        const int side = SizeGripSide(dlg);
        const RECT grip{ client.right - side, client.bottom - side, client.right, client.bottom };
        if (!EqualRect(&grip, &g_grip))
        {
            if (!IsRectEmpty(&g_grip)) InvalidateRect(dlg, &g_grip, TRUE);
            InvalidateRect(dlg, &grip, TRUE);
            g_grip = grip;
        }
    }

    // The size it opens at, centred on where the dialog manager put it.
    void SizeToDefault(HWND dlg)
    {
        RECT win, want{ 0, 0, Px(dlg, kClientW), Px(dlg, kClientH) };
        GetWindowRect(dlg, &win);
        typedef BOOL (WINAPI* AdjustFn)(LPRECT, DWORD, BOOL, DWORD, UINT);
        static const auto adjust = reinterpret_cast<AdjustFn>(reinterpret_cast<void*>(
            GetProcAddress(GetModuleHandleW(L"user32.dll"), "AdjustWindowRectExForDpi")));
        const DWORD style = static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_STYLE));
        const DWORD ex    = static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_EXSTYLE));
        if (adjust) adjust(&want, style, FALSE, ex, static_cast<UINT>(WindowDpi(dlg)));
        else        AdjustWindowRectEx(&want, style, FALSE, ex);
        const int w = want.right - want.left, h = want.bottom - want.top;
        SetWindowPos(dlg, nullptr, (win.left + win.right - w) / 2, (win.top + win.bottom - h) / 2,
                     w, h, SWP_NOZORDER | SWP_NOACTIVATE);
    }

    void ApplyDpi(HWND dlg)
    {
        MakeFonts(dlg);
        for (int id : kEditIds)
        {
            HWND e = GetDlgItem(dlg, id);
            SendMessageW(e, WM_SETFONT, reinterpret_cast<WPARAM>(FieldFont()), FALSE);
            SendMessageW(e, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, 0);
            SetWindowPos(e, nullptr, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER |
                                                 SWP_NOACTIVATE | SWP_FRAMECHANGED);
        }
        SendDlgItemMessageW(dlg, IDC_DIAG_CATEGORIES, LB_SETITEMHEIGHT, 0, Px(dlg, 27));
        grid::SetDpi(g_list, DpiOf(dlg));
        Layout(dlg);
        RedrawWindow(dlg, nullptr, nullptr, RDW_INVALIDATE | RDW_ERASE | RDW_FRAME | RDW_ALLCHILDREN);
    }

    void DrawPageIcon(const DRAWITEMSTRUCT* di)
    {
        HWND h = di->hwndItem;
        Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& r) {
            FillRect(dc, &r, PageBrush());
            glyph::Draw(dc, r, kPages[g_page].glyph, Line(h, 1), kIconInk, kPage); });
    }

    // ---- what the commands do --------------------------------------------------------------

    void Report(HWND dlg, trace::Result r)
    {
        if (r != trace::Result::Ok)
            MessageBoxW(dlg, trace::Explain(r), L"XRayXL Diagnostics", MB_OK | MB_ICONINFORMATION);
    }

    void ExportRows(HWND dlg, const diag::Table& rows)
    {
        if (rows.rows.empty()) return;
        const std::wstring path = diag::ExportPath(kPages[g_page].name);
        if (path.empty() || !diag::WriteBytes(path, diag::AsCsv(rows)))
        {
            MessageBoxW(dlg, L"The diagnostics file could not be written.\r\n\r\n"
                             L"XRayXL writes it beside the trace files, under %TEMP%\\XRayXL.",
                        L"XRayXL Diagnostics", MB_OK | MB_ICONINFORMATION);
            return;
        }
        core::Log::Note("diagnostics: exported");
        Report(dlg, trace::RevealFile(path));
    }

    void RevealChosen(HWND dlg)
    {
        const int col = kPages[g_page].pathColumn;
        const diag::Table picked = grid::Chosen(g_list);
        if (col < 0 || picked.rows.size() != 1) return;
        Report(dlg, trace::RevealFile(picked.rows[0][static_cast<size_t>(col)]));
    }

    void ShowListMenu(HWND dlg)
    {
        const int rows = grid::RowCount(g_list), picked = grid::ChosenCount(g_list);
        // Explorer selects one file at a time, so Reveal waits for one row.
        const bool canReveal = kPages[g_page].pathColumn >= 0 && picked == 1;

        HMENU menu = CreatePopupMenu();
        const auto add = [&](UINT cmd, const wchar_t* text, bool on) {
            AppendMenuW(menu, MF_STRING | (on ? 0u : MF_GRAYED), cmd, text); };
        add(IDM_DIAG_SELECTALL, L"Select &All\tCtrl+A", rows > 0);
        AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
        add(IDM_DIAG_REVEAL, L"&Reveal in File Explorer", canReveal);
        add(IDM_DIAG_COPY,   L"&Copy\tCtrl+C", picked > 0);
        add(IDM_DIAG_EXPORT, L"&Export to CSV", picked > 0);

        const POINT at = grid::MenuPoint(g_list);
        const int cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_RIGHTBUTTON, at.x, at.y, 0, dlg, nullptr);
        DestroyMenu(menu);

        if (cmd == IDM_DIAG_SELECTALL) grid::SelectAll(g_list);
        if (cmd == IDM_DIAG_REVEAL)    RevealChosen(dlg);
        if (cmd == IDM_DIAG_COPY)      Report(dlg, trace::CopyText(dlg, diag::AsTabbed(grid::Chosen(g_list))));
        if (cmd == IDM_DIAG_EXPORT)    ExportRows(dlg, grid::Chosen(g_list));
    }

    void InitControls(HWND dlg)
    {
        for (const PageDef& p : kPages)
            SendDlgItemMessageW(dlg, IDC_DIAG_CATEGORIES, LB_ADDSTRING, 0, reinterpret_cast<LPARAM>(p.name));
        SendDlgItemMessageW(dlg, IDC_DIAG_CATEGORIES, LB_SETCURSEL, 0, 0);

        for (HWND c = GetWindow(dlg, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT))
        {
            wchar_t cls[16] = {};
            GetClassNameW(c, cls, 16);
            const LONG_PTR st = GetWindowLongPtrW(c, GWL_STYLE);
            if (!lstrcmpiW(cls, L"Static") && (st & SS_TYPEMASK) == SS_LEFT)
                SetWindowLongPtrW(c, GWL_STYLE, (st & ~static_cast<LONG_PTR>(SS_TYPEMASK)) | SS_OWNERDRAW);
        }

        Subclass(GetDlgItem(dlg, IDC_DIAG_CATEGORIES), ListProc);
        SetWindowPos(GetDlgItem(dlg, IDC_DIAG_CATEGORIES), nullptr, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
        for (int id : kEditIds) Subclass(GetDlgItem(dlg, id), EditProc);

        g_list = grid::Create(dlg, IDC_DIAG_LIST);
        ShowWindow(g_list, SW_SHOW);
    }

    INT_PTR CALLBACK Proc(HWND dlg, UINT msg, WPARAM wp, LPARAM lp)
    {
        switch (msg)
        {
        case WM_INITDIALOG:
        {
            g_page = 0;
            ReadEverything();
            excelstyle::Begin();
            InitControls(dlg);
            excelstyle::WhiteCaption(dlg);

            typedef BOOL (WINAPI* BehaviourFn)(HWND, int, int);
            const auto behaviour = reinterpret_cast<BehaviourFn>(reinterpret_cast<void*>(
                GetProcAddress(GetModuleHandleW(L"user32.dll"), "SetDialogDpiChangeBehavior")));
            if (behaviour) behaviour(dlg, 7 /*DDC_DISABLE_ALL*/, 7);

            SizeToDefault(dlg);
            ApplyDpi(dlg);
            LoadPage(dlg, 0);
            SetFocus(g_list);
            return FALSE;                  // the list takes the focus, not the first tab stop
        }

        case WM_GETMINMAXINFO:
        {
            MINMAXINFO* mm = reinterpret_cast<MINMAXINFO*>(lp);
            if (!mm) return FALSE;
            RECT want{ 0, 0, Px(dlg, kMinW), Px(dlg, kMinH) };
            AdjustWindowRectEx(&want, static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_STYLE)), FALSE,
                               static_cast<DWORD>(GetWindowLongPtrW(dlg, GWL_EXSTYLE)));
            mm->ptMinTrackSize.x = want.right - want.left;
            mm->ptMinTrackSize.y = want.bottom - want.top;
            return TRUE;
        }

        case WM_SIZE:
            if (g_list) Layout(dlg);
            return TRUE;

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
            if (mi) mi->itemHeight = static_cast<UINT>(Px(dlg, 27));
            return TRUE;
        }

        case WM_DRAWITEM:
        {
            const DRAWITEMSTRUCT* di = reinterpret_cast<const DRAWITEMSTRUCT*>(lp);
            if (!di) return FALSE;
            const int id = static_cast<int>(di->CtlID);
            if (id == IDC_DIAG_CATEGORIES)    DrawCategory(di);
            else if (id == IDC_DIAG_PAGEICON) DrawPageIcon(di);
            else if (di->CtlType == ODT_STATIC)
                DrawLabel(di, In(kTitleIds, id) ? Face::Title : Face::Body,
                          ui::text::kVCentre | ui::text::kPrefix);
            else return FALSE;
            return TRUE;
        }

        case WM_NCHITTEST:
        {
            // The dots are an affordance, so the whole square they cover behaves like the
            // sizing corner. A dialog procedure returns its value through DWLP_MSGRESULT.
            POINT at{ GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            ScreenToClient(dlg, &at);
            RECT client;
            GetClientRect(dlg, &client);
            const int side = SizeGripSide(dlg);
            if (at.x >= client.right - side && at.y >= client.bottom - side &&
                at.x < client.right && at.y < client.bottom)
            {
                SetWindowLongPtrW(dlg, DWLP_MSGRESULT, HTBOTTOMRIGHT);
                return TRUE;
            }
            return FALSE;
        }

        case WM_PAINT:
        {
            // The background has already been erased by WM_CTLCOLORDLG's brush, and the
            // children are clipped out, so this only ever draws into the corner margin.
            PAINTSTRUCT ps;
            HDC dc = BeginPaint(dlg, &ps);
            RECT client;
            GetClientRect(dlg, &client);
            const int side = SizeGripSide(dlg);
            DrawSizeGrip(dlg, dc, RECT{ client.right - side, client.bottom - side,
                                        client.right, client.bottom });
            EndPaint(dlg, &ps);
            return TRUE;
        }

        case WM_CTLCOLORDLG:
            return reinterpret_cast<INT_PTR>(PageBrush());
        case WM_CTLCOLORLISTBOX:
        {
            const bool categories = reinterpret_cast<HWND>(lp) == GetDlgItem(dlg, IDC_DIAG_CATEGORIES);
            SetBkColor(reinterpret_cast<HDC>(wp), categories ? kPane : kPage);
            return reinterpret_cast<INT_PTR>(categories ? PaneBrush() : PageBrush());
        }
        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC:
        {
            HDC dc = reinterpret_cast<HDC>(wp);
            const bool lit = EditLit(reinterpret_cast<HWND>(lp));
            SetTextColor(dc, kText);
            SetBkColor(dc, lit ? kHotFace : kPage);
            return reinterpret_cast<INT_PTR>(lit ? HotBrush() : PageBrush());
        }

        case WM_COMMAND:
        {
            const int id = LOWORD(wp);
            const int code = HIWORD(wp);
            if (id == IDC_DIAG_CATEGORIES && code == LBN_SELCHANGE)
            {
                const int sel = static_cast<int>(SendDlgItemMessageW(dlg, IDC_DIAG_CATEGORIES, LB_GETCURSEL, 0, 0));
                if (sel >= 0 && sel < kPageCount) LoadPage(dlg, sel);
                return TRUE;
            }
            if (id == IDC_DIAG_SEARCH && code == EN_CHANGE)
            {
                wchar_t text[256] = {};
                GetDlgItemTextW(dlg, IDC_DIAG_SEARCH, text, 256);
                grid::SetFilter(g_list, text);
                return TRUE;
            }
            if (id == IDC_DIAG_LIST && code == grid::kContextMenu) { ShowListMenu(dlg); return TRUE; }
            if (id == IDCANCEL || id == IDOK) { EndDialog(dlg, IDCANCEL); return TRUE; }
            return FALSE;
        }

        case WM_CLOSE:
            EndDialog(dlg, IDCANCEL);
            return TRUE;

        case WM_DESTROY:
            g_list = nullptr;
            g_grip = RECT{};
            for (diag::Table& t : g_snapshot) t = diag::Table{};
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
        core::Log::Warning("diagnostics: Direct2D/DirectWrite unavailable; the dialog cannot draw its text");
        return false;
    }

    const excelstyle::DpiScope perMonitor;
    const INT_PTR r = DialogBoxParamW(excelstyle::Self(), MAKEINTRESOURCEW(IDD_XRAY_DIAG), owner, Proc, 0);
    grid::Unregister();          // no window of the class is left once the dialog has gone
    if (r == -1)
    {
        char line[64];
        _snprintf_s(line, _TRUNCATE, "diagnostics: dialog failed to open (%lu)", GetLastError());
        core::Log::Warning(line);
        return false;
    }
    core::Log::Note("diagnostics: closed");
    return true;
}
}
}
