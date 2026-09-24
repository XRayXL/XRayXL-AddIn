#include "gridlist.h"
#include "softtext.h"

#include <windowsx.h>

#include <algorithm>
#include <cstring>
#include <cwchar>
#include <cwctype>

// Every colour and distance here matches Excel's Name Manager at 96 DPI.

EXTERN_C IMAGE_DOS_HEADER __ImageBase;

namespace ui
{
namespace grid
{
namespace
{
    using ui::text::Face;

    constexpr COLORREF kPaper      = RGB(0xFF, 0xFF, 0xFF);
    constexpr COLORREF kFrame      = RGB(0x82, 0x87, 0x90);
    constexpr COLORREF kHeadRule   = RGB(0xE5, 0xE5, 0xE5);   // between header cells, and under them
    // Barely off the paper: any greyer and the band reads as scroll-bar furniture, which the
    // track beside it already is.
    constexpr COLORREF kHeadFill   = RGB(0xFD, 0xFD, 0xFD);
    constexpr COLORREF kHeadSorted = RGB(0xD9, 0xEB, 0xF9);
    constexpr COLORREF kSelection  = RGB(0x00, 0x78, 0xD7);
    constexpr COLORREF kFocusInk   = RGB(0xFF, 0x87, 0x28);
    constexpr COLORREF kText       = RGB(0x00, 0x00, 0x00);
    // A heading is a label, not data, so it reads lighter -- as File Explorer's do.
    constexpr COLORREF kHeadText   = RGB(0x5D, 0x5D, 0x5D);
    constexpr COLORREF kTextOn     = RGB(0xFF, 0xFF, 0xFF);
    constexpr COLORREF kTrack      = RGB(0xF0, 0xF0, 0xF0);   // the scroll bars' own track

    constexpr int kHeader96 = 25;
    constexpr int kRow96    = 17;
    constexpr int kPad96    = 7;     // a cell's text from its column's left edge
    constexpr int kGrip96   = 4;     // how near a separator counts as a drag

    const wchar_t* const kClass = L"XRayXLGridList";

    struct State
    {
        diag::Table       table;
        std::vector<int>  shown;          // indices into table.rows, filtered then sorted
        std::vector<int>  width;          // per column, in pixels at this DPI
        std::wstring      filter;
        int  sortColumn = -1;
        bool sortDown   = false;
        std::vector<char> chosen;         // one per entry in `shown`
        int  focus      = -1;             // index into `shown`: where the keyboard is
        int  anchor     = -1;             // where a Shift range starts
        int  top        = 0;              // first row drawn
        int  scrollX    = 0;
        int  dpi        = 96;
        bool focused    = false;
        int  dragColumn = -1;             // the separator being dragged
        int  dragFromX  = 0, dragFromWidth = 0;
        POINT menuAt{};                   // where the last context menu was asked for
    };

    State* Get(HWND h) { return reinterpret_cast<State*>(GetWindowLongPtrW(h, GWLP_USERDATA)); }

    int Px(const State& s, int at96) { return MulDiv(at96, s.dpi, 96); }
    int RowHeight(const State& s)    { return Px(s, kRow96); }
    int HeaderHeight(const State& s) { return Px(s, kHeader96); }

    int TotalWidth(const State& s)
    {
        int w = 0;
        for (int c : s.width) w += c;
        return w;
    }

    // ---- filtering and sorting ---------------------------------------------------------

    bool Holds(const std::wstring& hay, const std::wstring& needle)
    {
        if (needle.empty()) return true;
        if (needle.size() > hay.size()) return false;
        const size_t last = hay.size() - needle.size();
        for (size_t at = 0; at <= last; ++at)
        {
            size_t i = 0;
            while (i < needle.size() && towlower(hay[at + i]) == towlower(needle[i])) ++i;
            if (i == needle.size()) return true;
        }
        return false;
    }

    // A grouped byte count sorts by size, not by digit; anything else sorts as text.
    bool AsNumber(const std::wstring& s, double& out)
    {
        std::wstring plain;
        for (wchar_t c : s)
        {
            if (c == L',' || c == L' ') continue;
            if (!iswdigit(c) && c != L'.' && c != L'-') return false;
            plain += c;
        }
        if (plain.empty()) return false;
        wchar_t* end = nullptr;
        out = wcstod(plain.c_str(), &end);
        return end && *end == 0;
    }

    // The first cell identifies a row well enough to keep it selected across a sort or a filter.
    std::wstring KeyOf(const State& s, int shownIndex)
    {
        if (shownIndex < 0 || shownIndex >= static_cast<int>(s.shown.size())) return L"";
        const std::vector<std::wstring>& row = s.table.rows[s.shown[shownIndex]];
        return row.empty() ? std::wstring() : row[0];
    }

    void ClearChosen(State& s)
    {
        s.chosen.assign(s.shown.size(), 0);
    }

    void Rebuild(State& s)
    {
        const std::wstring keep = KeyOf(s, s.focus);

        s.shown.clear();
        for (size_t r = 0; r < s.table.rows.size(); ++r)
        {
            bool hit = s.filter.empty();
            for (size_t c = 0; !hit && c < s.table.rows[r].size(); ++c)
                hit = Holds(s.table.rows[r][c], s.filter);
            if (hit) s.shown.push_back(static_cast<int>(r));
        }

        if (s.sortColumn >= 0 && s.sortColumn < static_cast<int>(s.table.columns.size()))
        {
            const int col = s.sortColumn;
            const bool down = s.sortDown;
            const std::vector<std::vector<std::wstring>>& rows = s.table.rows;
            std::stable_sort(s.shown.begin(), s.shown.end(), [&](int a, int b) {
                const std::wstring& x = rows[a][col];
                const std::wstring& y = rows[b][col];
                double nx = 0, ny = 0;
                const int cmp = (AsNumber(x, nx) && AsNumber(y, ny))
                              ? (nx < ny ? -1 : nx > ny ? 1 : 0)
                              : _wcsicmp(x.c_str(), y.c_str());
                return down ? cmp > 0 : cmp < 0;
            });
        }

        ClearChosen(s);
        s.focus = s.anchor = -1;
        if (!keep.empty())
            for (size_t i = 0; i < s.shown.size(); ++i)
            {
                const std::vector<std::wstring>& row = s.table.rows[s.shown[i]];
                if (!row.empty() && row[0] == keep)
                {
                    s.focus = s.anchor = static_cast<int>(i);
                    s.chosen[i] = 1;
                    break;
                }
            }
        s.top = 0;
    }

    // ---- scrolling ---------------------------------------------------------------------

    // Rows that fit whole: what a page of scrolling moves by, and what the scroll bar's page
    // size means.
    int PageRows(HWND h, const State& s)
    {
        RECT c;
        GetClientRect(h, &c);
        return (std::max)(1, (static_cast<int>(c.bottom) - HeaderHeight(s)) / RowHeight(s));
    }

    // Rows that show any of themselves. A row whose bottom is cut off is still drawn: leaving
    // it out tells the reader the data ends there.
    int DrawnRows(HWND h, const State& s)
    {
        RECT c;
        GetClientRect(h, &c);
        const int body = (std::max)(0, static_cast<int>(c.bottom) - HeaderHeight(s));
        return (body + RowHeight(s) - 1) / RowHeight(s);
    }

    // The last column takes whatever the others leave, so a narrow table shows no scroll bar.
    // A drag on a column overrides it until the next resize.
    void FitLastColumn(HWND h, State& s)
    {
        if (s.width.empty()) return;
        RECT c;
        GetClientRect(h, &c);
        const int visible = static_cast<int>(c.right);
        int others = 0;
        for (size_t i = 0; i + 1 < s.width.size(); ++i) others += s.width[i];
        const int natural = MulDiv(s.table.columns.back().width96, s.dpi, 96);
        s.width.back() = (std::max)(natural, visible - others);
    }

    void SetScrollBarsOnce(HWND h, State& s)
    {
        RECT c;
        GetClientRect(h, &c);
        const int page = PageRows(h, s);
        const int rows = static_cast<int>(s.shown.size());

        s.top = (std::max)(0, (std::min)(s.top, (std::max)(0, rows - page)));
        SCROLLINFO v{ sizeof(v), SIF_RANGE | SIF_PAGE | SIF_POS, 0, (std::max)(0, rows - 1),
                      static_cast<UINT>(page), s.top, 0 };
        SetScrollInfo(h, SB_VERT, &v, TRUE);

        const int visible = (std::max)(1, static_cast<int>(c.right));
        const int total = TotalWidth(s);
        s.scrollX = (std::max)(0, (std::min)(s.scrollX, (std::max)(0, total - visible)));
        SCROLLINFO hz{ sizeof(hz), SIF_RANGE | SIF_PAGE | SIF_POS, 0, (std::max)(0, total - 1),
                       static_cast<UINT>(visible), s.scrollX, 0 };
        SetScrollInfo(h, SB_HORZ, &hz, TRUE);
    }

    // Showing or hiding one bar changes the client the other is measured against, so the pair
    // is set twice -- never in a loop, which could oscillate.
    void SyncScrollBars(HWND h, State& s)
    {
        RECT before;
        GetClientRect(h, &before);
        SetScrollBarsOnce(h, s);
        RECT after;
        GetClientRect(h, &after);
        if (after.right != before.right || after.bottom != before.bottom) SetScrollBarsOnce(h, s);
    }

    void ShowRow(HWND h, State& s, int row)
    {
        const int page = PageRows(h, s);
        if (row < s.top) s.top = row;
        else if (row >= s.top + page) s.top = row - page + 1;
        SyncScrollBars(h, s);
    }

    void Announce(HWND h, WORD what)
    {
        SendMessageW(GetParent(h), WM_COMMAND,
                     MAKEWPARAM(GetDlgCtrlID(h), what), reinterpret_cast<LPARAM>(h));
    }

    void ChooseRange(State& s, int from, int to)
    {
        if (from < 0 || to < 0) return;
        if (from > to) std::swap(from, to);
        for (int i = from; i <= to && i < static_cast<int>(s.chosen.size()); ++i) s.chosen[i] = 1;
    }

    // `extend` is Shift: the range from the anchor. `toggle` is Ctrl: this row alone, flipped.
    void Choose(HWND h, State& s, int row, bool extend, bool toggle, bool scroll)
    {
        const int rows = static_cast<int>(s.shown.size());
        if (rows == 0) { s.focus = s.anchor = -1; ClearChosen(s); InvalidateRect(h, nullptr, FALSE); return; }
        row = (std::max)(0, (std::min)(row, rows - 1));

        if (extend && s.anchor >= 0)
        {
            ClearChosen(s);
            ChooseRange(s, s.anchor, row);
        }
        else if (toggle)
        {
            if (static_cast<int>(s.chosen.size()) > row) s.chosen[row] = s.chosen[row] ? 0 : 1;
            s.anchor = row;
        }
        else
        {
            ClearChosen(s);
            if (static_cast<int>(s.chosen.size()) > row) s.chosen[row] = 1;
            s.anchor = row;
        }
        s.focus = row;
        if (scroll) ShowRow(h, s, row);
        InvalidateRect(h, nullptr, FALSE);
    }

    // ---- painting ----------------------------------------------------------------------

    void Fill(HDC dc, const RECT& r, COLORREF colour)
    {
        HBRUSH b = CreateSolidBrush(colour);
        FillRect(dc, &r, b);
        DeleteObject(b);
    }

    // Excel's focus ring: one pixel on, one off, all the way round the row.
    void DottedRect(HDC dc, const RECT& r, COLORREF colour, int step)
    {
        for (int x = r.left; x < r.right; x += step * 2)
        {
            SetPixelV(dc, x, r.top, colour);
            SetPixelV(dc, x, r.bottom - 1, colour);
        }
        for (int y = r.top; y < r.bottom; y += step * 2)
        {
            SetPixelV(dc, r.left, y, colour);
            SetPixelV(dc, r.right - 1, y, colour);
        }
    }

    // Clipped to its own column, so a long path stops at the next one rather than running over it.
    void PutCellText(ui::text::Batch& batch, const State& s, const RECT& cell, const RECT& clip,
                     const std::wstring& text, bool rightAlign, COLORREF ink, Face face)
    {
        if (text.empty()) return;
        RECT t = cell;
        t.left += Px(s, kPad96);
        t.right -= Px(s, kPad96);
        if (t.right <= t.left) return;
        if (rightAlign)
        {
            SIZE sz{};
            if (ui::text::Measure(text.c_str(), face, 0, s.dpi, sz) && sz.cx < t.right - t.left)
                t.left = t.right - sz.cx;
        }
        const RECT box{ (std::max)(clip.left, cell.left), clip.top,
                        (std::min)(clip.right, cell.right), clip.bottom };
        batch.Put(t, box, text.c_str(), face, ink, ui::text::kVCentre, s.dpi);
    }

    void Paint(HWND h, HDC target)
    {
        State* sp = Get(h);
        if (!sp) return;
        State& s = *sp;
        RECT c;
        GetClientRect(h, &c);
        const int w = c.right, ht = c.bottom;
        if (w <= 0 || ht <= 0) return;

        HDC mem = CreateCompatibleDC(target);
        HBITMAP bmp = CreateCompatibleBitmap(target, w, ht);
        if (!mem || !bmp)
        {
            if (bmp) DeleteObject(bmp);
            if (mem) DeleteDC(mem);
            return;
        }
        HGDIOBJ old = SelectObject(mem, bmp);

        const int line = (std::max)(1, s.dpi / 96);
        const int headH = HeaderHeight(s), rowH = RowHeight(s);

        Fill(mem, c, kPaper);

        // the frame is the non-client edge, so the scroll bars fall inside it
        const int inner = 0, innerRight = c.right;
        const int headTop = 0, bodyTop = headH;

        // Every GDI stroke first, then the text in one batch: a batch reaches the DC only
        // when it closes, so anything drawn under the text has to be there already.
        const int page = DrawnRows(h, s);
        const int rows = static_cast<int>(s.shown.size());

        // A hint of grey behind the headings, and a rule under the band, so it reads as a strip
        // of labels rather than a first row of data.
        Fill(mem, RECT{ inner, headTop + line, innerRight, bodyTop }, kHeadFill);

        int x = inner - s.scrollX;
        for (size_t col = 0; col < s.table.columns.size(); ++col)
        {
            const int right = x + s.width[col];
            const RECT cell{ (std::max)(x, inner), headTop + line, (std::min)(right, innerRight), bodyTop };
            if (cell.right > cell.left && static_cast<int>(col) == s.sortColumn)
                Fill(mem, cell, kHeadSorted);
            if (right > inner && right < innerRight)
                Fill(mem, RECT{ right - line, headTop, right, bodyTop }, kHeadRule);
            x = right;
        }
        Fill(mem, RECT{ inner, bodyTop - line, innerRight, bodyTop }, kHeadRule);
        for (int i = 0; i < page && s.top + i < rows; ++i)
        {
            const int row = s.top + i;
            // `full` is the row; `shown` is as much of it as fits. The last row may be cut off,
            // and everything positional has to come from `full` or it moves as the window does.
            const RECT full{ inner, bodyTop + i * rowH, innerRight, bodyTop + (i + 1) * rowH };
            if (full.top >= c.bottom) break;
            RECT shown = full;
            shown.bottom = (std::min)(shown.bottom, c.bottom);
            if (row < static_cast<int>(s.chosen.size()) && s.chosen[row])
                Fill(mem, shown, kSelection);
            // on `full`: a cut-off row's ring runs off the bottom rather than closing early
            if (row == s.focus && s.focused) DottedRect(mem, full, kFocusInk, line);
        }

        {
            ui::text::Batch batch(mem, c);
            x = inner - s.scrollX;
            for (size_t col = 0; col < s.table.columns.size(); ++col)
            {
                const int right = x + s.width[col];
                if (right > inner && x < innerRight)
                    // left aligned whatever the column's data does, as File Explorer's are
                    PutCellText(batch, s, RECT{ x, headTop + line, right, bodyTop },
                                RECT{ inner, headTop + line, innerRight, bodyTop },
                                s.table.columns[col].title, false, kHeadText, Face::Body);
                x = right;
            }
            for (int i = 0; i < page && s.top + i < rows; ++i)
            {
                const int row = s.top + i;
                const RECT full{ inner, bodyTop + i * rowH, innerRight, bodyTop + (i + 1) * rowH };
                if (full.top >= c.bottom) break;
                RECT shown = full;
                shown.bottom = (std::min)(shown.bottom, c.bottom);
                const bool on = (row < static_cast<int>(s.chosen.size())) && s.chosen[row] != 0;
                const std::vector<std::wstring>& cells = s.table.rows[s.shown[row]];
                int cx = inner - s.scrollX;
                for (size_t col = 0; col < s.table.columns.size(); ++col)
                {
                    const int right = cx + s.width[col];
                    // laid out in the whole row so it sits where every other row's does, and
                    // clipped to what fits so a cut-off row is cut off rather than moved
                    if (right > inner && cx < innerRight && col < cells.size())
                        PutCellText(batch, s, RECT{ cx, full.top, right, full.bottom },
                                    RECT{ inner, shown.top, innerRight, shown.bottom },
                                    cells[col], s.table.columns[col].rightAlign,
                                    on ? kTextOn : kText, Face::Body);
                    cx = right;
                }
            }
        }

        BitBlt(target, 0, 0, w, ht, mem, 0, 0, SRCCOPY);
        SelectObject(mem, old);
        DeleteObject(bmp);
        DeleteDC(mem);
    }

    // ---- hit testing --------------------------------------------------------------------

    // The column whose right-hand separator is within a grip of `x`, or -1.
    int SeparatorAt(const State& s, int x, int inner)
    {
        int at = inner - s.scrollX;
        for (size_t col = 0; col < s.width.size(); ++col)
        {
            at += s.width[col];
            if (x >= at - Px(s, kGrip96) && x <= at + Px(s, kGrip96)) return static_cast<int>(col);
        }
        return -1;
    }

    int ColumnAt(const State& s, int x, int inner)
    {
        int at = inner - s.scrollX;
        for (size_t col = 0; col < s.width.size(); ++col)
        {
            if (x >= at && x < at + s.width[col]) return static_cast<int>(col);
            at += s.width[col];
        }
        return -1;
    }

    int RowAt(const State& s, int y)
    {
        const int bodyTop = HeaderHeight(s);
        if (y < bodyTop) return -1;
        const int row = s.top + (y - bodyTop) / RowHeight(s);
        return (row < static_cast<int>(s.shown.size())) ? row : -1;
    }

    void SortBy(HWND h, State& s, int column)
    {
        if (column < 0) return;
        if (s.sortColumn == column) s.sortDown = !s.sortDown;
        else { s.sortColumn = column; s.sortDown = false; }
        Rebuild(s);
        SyncScrollBars(h, s);
        if (s.focus >= 0) ShowRow(h, s, s.focus);
        InvalidateRect(h, nullptr, FALSE);
    }

    // Tab separated with a header, which is what Excel pastes into cells.
    void CopyChosen(HWND h, const State& s)
    {
        std::wstring text;
        for (size_t c = 0; c < s.table.columns.size(); ++c)
        {
            if (c) text += L'\t';
            text += s.table.columns[c].title;
        }
        text += L"\r\n";
        for (size_t i = 0; i < s.shown.size(); ++i)
        {
            if (i < s.chosen.size() && !s.chosen[i]) continue;
            const std::vector<std::wstring>& row = s.table.rows[s.shown[i]];
            for (size_t c = 0; c < s.table.columns.size(); ++c)
            {
                if (c) text += L'\t';
                if (c < row.size()) text += row[c];
            }
            text += L"\r\n";
        }
        if (!OpenClipboard(h)) return;
        EmptyClipboard();
        const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
        if (HGLOBAL mem = GlobalAlloc(GMEM_MOVEABLE, bytes))
        {
            if (void* at = GlobalLock(mem))
            {
                memcpy(at, text.c_str(), bytes);
                GlobalUnlock(mem);
                SetClipboardData(CF_UNICODETEXT, mem);
            }
            else GlobalFree(mem);
        }
        CloseClipboard();
    }

    LRESULT CALLBACK Proc(HWND h, UINT msg, WPARAM wp, LPARAM lp)
    {
        State* s = Get(h);
        switch (msg)
        {
        case WM_NCCREATE:
            SetWindowLongPtrW(h, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(new State()));
            return TRUE;

        case WM_NCDESTROY:
            delete s;
            SetWindowLongPtrW(h, GWLP_USERDATA, 0);
            return 0;

        case WM_GETDLGCODE:
            return DLGC_WANTARROWS | DLGC_WANTCHARS;

        case WM_NCCALCSIZE:
        {
            // Reserved before DefWindowProc places the scroll bars, so they land inside the
            // frame rather than beside it. wParam says which shape lParam is.
            if (s && lp)
            {
                const int line = (std::max)(1, s->dpi / 96);
                RECT* r = wp ? &reinterpret_cast<NCCALCSIZE_PARAMS*>(lp)->rgrc[0]
                             : reinterpret_cast<RECT*>(lp);
                InflateRect(r, -line, -line);
            }
            return DefWindowProcW(h, msg, wp, lp);
        }

        case WM_NCPAINT:
        {
            const LRESULT r = DefWindowProcW(h, msg, wp, lp);     // the scroll bars
            if (!s) return r;
            RECT win;
            GetWindowRect(h, &win);
            const POINT origin{ win.left, win.top };
            RECT w = win;
            OffsetRect(&w, -origin.x, -origin.y);
            const int line = (std::max)(1, s->dpi / 96);
            if (HDC dc = GetWindowDC(h))
            {
                // Whether DefWindowProc fills the square between the two scroll bars is not
                // something to depend on: it is filled here, so it looks the same every time.
                SCROLLBARINFO v{}, z{};
                v.cbSize = z.cbSize = sizeof(SCROLLBARINFO);
                const bool haveV = GetScrollBarInfo(h, OBJID_VSCROLL, &v) &&
                                   !(v.rgstate[0] & STATE_SYSTEM_INVISIBLE);
                const bool haveH = GetScrollBarInfo(h, OBJID_HSCROLL, &z) &&
                                   !(z.rgstate[0] & STATE_SYSTEM_INVISIBLE);
                if (haveV && haveH)
                {
                    RECT corner{ v.rcScrollBar.left - origin.x, z.rcScrollBar.top - origin.y,
                                 v.rcScrollBar.right - origin.x, z.rcScrollBar.bottom - origin.y };
                    if (corner.right > corner.left && corner.bottom > corner.top)
                        Fill(dc, corner, kTrack);
                }
                Fill(dc, RECT{ w.left, w.top, w.right, w.top + line }, kFrame);
                Fill(dc, RECT{ w.left, w.bottom - line, w.right, w.bottom }, kFrame);
                Fill(dc, RECT{ w.left, w.top, w.left + line, w.bottom }, kFrame);
                Fill(dc, RECT{ w.right - line, w.top, w.right, w.bottom }, kFrame);
                ReleaseDC(h, dc);
            }
            return r;
        }

        case WM_ERASEBKGND:
            return 1;

        case WM_PAINT:
        {
            PAINTSTRUCT ps;
            Paint(h, BeginPaint(h, &ps));
            EndPaint(h, &ps);
            return 0;
        }

        case WM_SIZE:
            if (s)
            {
                FitLastColumn(h, *s);
                SyncScrollBars(h, *s);
                InvalidateRect(h, nullptr, FALSE);   // the last column moved; so did every row
            }
            return 0;

        case WM_SETFOCUS:
        case WM_KILLFOCUS:
            if (s) { s->focused = (msg == WM_SETFOCUS); InvalidateRect(h, nullptr, FALSE); }
            return 0;

        case WM_LBUTTONDOWN:
        {
            if (!s) break;
            SetFocus(h);
            const int x = GET_X_LPARAM(lp), y = GET_Y_LPARAM(lp);
            if (y < HeaderHeight(*s))
            {
                const int sep = SeparatorAt(*s, x, 0);
                if (sep >= 0)
                {
                    s->dragColumn = sep;
                    s->dragFromX = x;
                    s->dragFromWidth = s->width[sep];
                    SetCapture(h);
                }
                else SortBy(h, *s, ColumnAt(*s, x, 0));
                return 0;
            }
            // the message carries the modifiers; GetKeyState would answer for whichever
            // thread happens to be asking
            const int row = RowAt(*s, y);
            if (row >= 0)
                Choose(h, *s, row, (wp & MK_SHIFT) != 0, (wp & MK_CONTROL) != 0, false);
            return 0;
        }

        case WM_MOUSEMOVE:
            if (s && s->dragColumn >= 0 && s->dragColumn < static_cast<int>(s->width.size()))
            {
                const int want = s->dragFromWidth + (GET_X_LPARAM(lp) - s->dragFromX);
                s->width[s->dragColumn] = (std::max)(MulDiv(24, s->dpi, 96), want);
                SyncScrollBars(h, *s);
                InvalidateRect(h, nullptr, FALSE);
            }
            return 0;

        case WM_LBUTTONUP:
            if (s && s->dragColumn >= 0) { s->dragColumn = -1; ReleaseCapture(); }
            return 0;

        // Alt+Tab or anything else can take the mouse mid-drag, and no button-up follows.
        case WM_CAPTURECHANGED:
            if (s) s->dragColumn = -1;
            return 0;

        case WM_SETCURSOR:
            if (s && LOWORD(lp) == HTCLIENT)
            {
                POINT p;
                GetCursorPos(&p);
                ScreenToClient(h, &p);
                if (s->dragColumn >= 0 ||
                    (p.y < HeaderHeight(*s) && SeparatorAt(*s, p.x, 0) >= 0))
                {
                    SetCursor(LoadCursorW(nullptr, IDC_SIZEWE));
                    return TRUE;
                }
            }
            break;

        case WM_RBUTTONDOWN:
        {
            if (!s) break;
            SetFocus(h);
            const int row = RowAt(*s, GET_Y_LPARAM(lp));
            const bool already = row >= 0 && row < static_cast<int>(s->chosen.size()) && s->chosen[row];
            if (row >= 0 && !already) Choose(h, *s, row, false, false, false);
            return 0;
        }

        case WM_CONTEXTMENU:
        {
            if (!s) break;
            POINT at{ GET_X_LPARAM(lp), GET_Y_LPARAM(lp) };
            if (at.x == -1 && at.y == -1)          // from the keyboard: beside the focus row
            {
                const int row = (s->focus >= 0) ? s->focus - s->top : 0;
                at = POINT{ Px(*s, 24), HeaderHeight(*s) + (row + 1) * RowHeight(*s) };
                ClientToScreen(h, &at);
            }
            s->menuAt = at;
            Announce(h, kContextMenu);
            return 0;
        }

        case WM_VSCROLL:
        {
            if (!s) break;
            const int page = PageRows(h, *s);
            const int was = s->top;
            switch (LOWORD(wp))
            {
            case SB_LINEUP:   s->top -= 1; break;
            case SB_LINEDOWN: s->top += 1; break;
            case SB_PAGEUP:   s->top -= page; break;
            case SB_PAGEDOWN: s->top += page; break;
            case SB_TOP:      s->top = 0; break;
            case SB_BOTTOM:   s->top = static_cast<int>(s->shown.size()); break;
            case SB_THUMBPOSITION:
            case SB_THUMBTRACK:
            {
                SCROLLINFO si{ sizeof(si), SIF_TRACKPOS };
                if (GetScrollInfo(h, SB_VERT, &si)) s->top = si.nTrackPos;
                break;
            }
            default: break;
            }
            SyncScrollBars(h, *s);
            if (s->top != was) InvalidateRect(h, nullptr, FALSE);
            return 0;
        }

        case WM_HSCROLL:
        {
            if (!s) break;
            RECT c;
            GetClientRect(h, &c);
            const int step = MulDiv(24, s->dpi, 96), was = s->scrollX;
            switch (LOWORD(wp))
            {
            case SB_LINELEFT:  s->scrollX -= step; break;
            case SB_LINERIGHT: s->scrollX += step; break;
            case SB_PAGELEFT:  s->scrollX -= c.right; break;
            case SB_PAGERIGHT: s->scrollX += c.right; break;
            case SB_THUMBPOSITION:
            case SB_THUMBTRACK:
            {
                SCROLLINFO si{ sizeof(si), SIF_TRACKPOS };
                if (GetScrollInfo(h, SB_HORZ, &si)) s->scrollX = si.nTrackPos;
                break;
            }
            default: break;
            }
            SyncScrollBars(h, *s);
            if (s->scrollX != was) InvalidateRect(h, nullptr, FALSE);
            return 0;
        }

        case WM_MOUSEWHEEL:
        {
            if (!s) break;
            UINT lines = 3;
            SystemParametersInfoW(SPI_GETWHEELSCROLLLINES, 0, &lines, 0);
            s->top -= GET_WHEEL_DELTA_WPARAM(wp) * static_cast<int>(lines) / WHEEL_DELTA;
            SyncScrollBars(h, *s);
            InvalidateRect(h, nullptr, FALSE);
            return 0;
        }

        case WM_KEYDOWN:
        {
            if (!s) break;
            const int page = PageRows(h, *s);
            const bool shift = GetKeyState(VK_SHIFT) < 0;
            const int at = (s->focus >= 0) ? s->focus : 0;
            switch (wp)
            {
            case VK_UP:     Choose(h, *s, at - 1, shift, false, true); return 0;
            case VK_DOWN:   Choose(h, *s, at + 1, shift, false, true); return 0;
            case VK_PRIOR:  Choose(h, *s, at - page, shift, false, true); return 0;
            case VK_NEXT:   Choose(h, *s, at + page, shift, false, true); return 0;
            case VK_HOME:   Choose(h, *s, 0, shift, false, true); return 0;
            case VK_END:    Choose(h, *s, static_cast<int>(s->shown.size()) - 1, shift, false, true); return 0;
            case 'A':
                if (GetKeyState(VK_CONTROL) < 0) { SelectAll(h); return 0; }
                break;
            case 'C':
                if (GetKeyState(VK_CONTROL) < 0) { CopyChosen(h, *s); return 0; }
                break;
            default: break;
            }
            break;
        }
        }
        return DefWindowProcW(h, msg, wp, lp);
    }

    void EnsureClass()
    {
        WNDCLASSEXW existing{};
        existing.cbSize = sizeof(existing);
        if (GetClassInfoExW(reinterpret_cast<HINSTANCE>(&__ImageBase), kClass, &existing)) return;
        WNDCLASSEXW wc{};
        wc.cbSize = sizeof(wc);
        wc.style = CS_DBLCLKS;
        wc.lpfnWndProc = Proc;
        wc.hInstance = reinterpret_cast<HINSTANCE>(&__ImageBase);
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        wc.lpszClassName = kClass;
        RegisterClassExW(&wc);
    }
}

HWND Create(HWND parent, int id)
{
    EnsureClass();
    return CreateWindowExW(0, kClass, L"", WS_CHILD | WS_TABSTOP | WS_VSCROLL | WS_HSCROLL,
                           0, 0, 10, 10, parent, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)),
                           reinterpret_cast<HINSTANCE>(&__ImageBase), nullptr);
}

void Unregister()
{
    UnregisterClassW(kClass, reinterpret_cast<HINSTANCE>(&__ImageBase));
}

void SetTable(HWND list, diag::Table table)
{
    State* s = Get(list);
    if (!s) return;
    // a drag in progress was over the old columns
    if (s->dragColumn >= 0) { s->dragColumn = -1; ReleaseCapture(); }
    s->table = std::move(table);
    s->width.clear();
    for (const diag::Column& c : s->table.columns) s->width.push_back(MulDiv(c.width96, s->dpi, 96));
    s->sortColumn = -1;
    s->sortDown = false;
    s->focus = s->anchor = -1;
    Rebuild(*s);
    FitLastColumn(list, *s);
    SyncScrollBars(list, *s);
    InvalidateRect(list, nullptr, FALSE);
}

void SetFilter(HWND list, const wchar_t* text)
{
    State* s = Get(list);
    if (!s) return;
    const std::wstring want = text ? text : L"";
    if (want == s->filter) return;
    s->filter = want;
    Rebuild(*s);
    FitLastColumn(list, *s);      // the row count changed, so the vertical bar may have gone
    SyncScrollBars(list, *s);
    InvalidateRect(list, nullptr, FALSE);
}

diag::Table Shown(HWND list)
{
    diag::Table out;
    State* s = Get(list);
    if (!s) return out;
    out.columns = s->table.columns;
    for (int i : s->shown) out.rows.push_back(s->table.rows[i]);
    return out;
}

int RowCount(HWND list)
{
    State* s = Get(list);
    return s ? static_cast<int>(s->shown.size()) : 0;
}

diag::Table Chosen(HWND list)
{
    diag::Table out;
    State* s = Get(list);
    if (!s) return out;
    out.columns = s->table.columns;
    for (size_t i = 0; i < s->shown.size(); ++i)
        if (i < s->chosen.size() && s->chosen[i]) out.rows.push_back(s->table.rows[s->shown[i]]);
    return out;
}

int ChosenCount(HWND list)
{
    State* s = Get(list);
    if (!s) return 0;
    int n = 0;
    for (char c : s->chosen) if (c) ++n;
    return n;
}

void SelectAll(HWND list)
{
    State* s = Get(list);
    if (!s || s->shown.empty()) return;
    s->chosen.assign(s->shown.size(), 1);
    if (s->focus < 0) s->focus = 0;
    s->anchor = 0;
    InvalidateRect(list, nullptr, FALSE);
}

int Focused(HWND list)
{
    State* s = Get(list);
    return s ? s->focus : -1;
}

POINT MenuPoint(HWND list)
{
    State* s = Get(list);
    return s ? s->menuAt : POINT{};
}

void SetDpi(HWND list, int dpi)
{
    State* s = Get(list);
    if (!s || dpi <= 0) return;
    const int was = s->dpi;
    s->dpi = dpi;
    for (int& w : s->width) w = MulDiv(w, dpi, was);
    FitLastColumn(list, *s);
    SyncScrollBars(list, *s);
    InvalidateRect(list, nullptr, FALSE);
}
}
}
