#include "eventpane.h"
#include "dlgexcelstyle.h"
#include "app/appevents.h"
#include "core/text.h"

#include <windows.h>
#include <commctrl.h>
#include <initguid.h>   // defines the accessibility GUIDs oleacc.h declares
#include <oleacc.h>
#include <cwchar>
#include <vector>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "oleacc.lib")

namespace ui
{
namespace eventpane
{
namespace
{
    using namespace ui::excelstyle;
    using ui::text::Face;
    namespace ev = core::events;

    constexpr int kGroups = static_cast<int>(ev::Group::Count_);
    const wchar_t kClass[] = L"XRayXLEventPane";

    // The drop-down's items, in this order.
    const ev::Preset kPresets[] = { ev::Preset::None, ev::Preset::Calc, ev::Preset::Selection,
                                    ev::Preset::CalcAndSelection, ev::Preset::All, ev::Preset::Custom };
    constexpr int kPresetCount = static_cast<int>(sizeof(kPresets) / sizeof(kPresets[0]));

    struct State
    {
        ev::Mask selected = 0;
        ev::Mask available = 0;
        // A group's mix before its heading was clicked, restored by the third click.
        ev::Mask remembered[kGroups] = {};
        bool     remembers[kGroups] = {};
        bool     enabled = true;
        int      scroll = 0;            // pixels scrolled
        int      content = 0;           // total height of the rows
        HWND     tip = nullptr;
        IAccPropServices* acc = nullptr;
        std::vector<HWND> boxes;        // headings then events, in drawing order
        std::vector<int>  shown;        // each box's state as last drawn, -1 before the first
    };

    State* Of(HWND pane) { return reinterpret_cast<State*>(GetWindowLongPtrW(pane, GWLP_USERDATA)); }

    bool IsGroupId(int id) { return id >= kFirstGroupId && id < kFirstGroupId + kGroups; }
    bool IsEventId(int id) { return id >= kFirstEventId && id < kFirstEventId + ev::Count(); }

    ev::Mask GroupMask(int g)
    {
        ev::Mask m = 0;
        for (int k = 0; k < ev::Count(); ++k)
            if (static_cast<int>(ev::At(k).group) == g) m |= ev::Mask(1) << k;
        return m;
    }

    int StateOf(const State& s, int g)
    {
        const ev::Mask g2 = GroupMask(g) & s.available;
        const ev::Mask on = s.selected & g2;
        if (!g2 || !on) return 0;
        return on == g2 ? 1 : 2;
    }

    int Count(ev::Mask m) { int n = 0; for (; m; m &= m - 1) ++n; return n; }

    // A heading's next click: mixed -> all, all -> none, none -> the remembered mix if there is one.
    void CycleGroup(State& s, int g)
    {
        const ev::Mask g2 = GroupMask(g) & s.available;
        switch (StateOf(s, g))
        {
        case 2:
            s.remembered[g] = s.selected & g2;
            s.remembers[g] = true;
            s.selected |= g2;
            break;
        case 1:
            s.selected &= ~g2;
            break;
        default:
            if (s.remembers[g]) { s.selected = (s.selected & ~g2) | s.remembered[g]; s.remembers[g] = false; }
            else                s.selected |= g2;
            break;
        }
    }

    // The role and state a screen reader announces, since an owner-drawn button reports neither.
    void Annotate(State& s, HWND box, int checkState, bool unavailable)
    {
        if (!s.acc) return;
        VARIANT role; VariantInit(&role); role.vt = VT_I4; role.lVal = ROLE_SYSTEM_CHECKBUTTON;
        s.acc->SetHwndProp(box, static_cast<DWORD>(OBJID_CLIENT), CHILDID_SELF, PROPID_ACC_ROLE, role);
        VARIANT st; VariantInit(&st); st.vt = VT_I4;
        st.lVal = STATE_SYSTEM_FOCUSABLE | (checkState == 1 ? STATE_SYSTEM_CHECKED : 0) |
                  (checkState == 2 ? STATE_SYSTEM_MIXED : 0) | (unavailable ? STATE_SYSTEM_UNAVAILABLE : 0);
        s.acc->SetHwndProp(box, static_cast<DWORD>(OBJID_CLIENT), CHILDID_SELF, PROPID_ACC_STATE, st);
    }

    // Box states into each control, for drawing and for the accessibility annotation. Only a
    // box whose state changed is touched: a tick changes two of some sixty.
    void Refresh(HWND pane)
    {
        State& s = *Of(pane);
        s.shown.resize(s.boxes.size(), -1);
        for (size_t i = 0; i < s.boxes.size(); ++i)
        {
            HWND b = s.boxes[i];
            const int id = GetDlgCtrlID(b);
            int cs = 0; bool unavailable = false;
            if (IsGroupId(id))
            {
                // A heading with none of its events in this Excel has nothing to tick.
                cs = StateOf(s, id - kFirstGroupId);
                unavailable = !(GroupMask(id - kFirstGroupId) & s.available);
            }
            else
            {
                const int k = id - kFirstEventId;
                cs = ((s.selected >> k) & 1) ? 1 : 0;
                unavailable = !((s.available >> k) & 1);
            }
            const bool enabled = s.enabled && !unavailable;
            const int now = cs | (enabled ? 4 : 0);
            if (now == s.shown[i]) continue;
            s.shown[i] = now;
            SetWindowLongPtrW(b, GWLP_USERDATA, cs);
            if ((IsWindowEnabled(b) != FALSE) != enabled) EnableWindow(b, enabled ? TRUE : FALSE);
            Annotate(s, b, cs, !enabled);
            InvalidateRect(b, nullptr, FALSE);
        }
    }

    int RowHeight(HWND pane, int id) { return Px(pane, IsGroupId(id) ? 28 : 22); }

    // Rows in content coordinates, moved by the scroll.
    void Layout(HWND pane)
    {
        State& s = *Of(pane);
        RECT rc; GetClientRect(pane, &rc);
        int y = 0;
        HDWP dwp = BeginDeferWindowPos(static_cast<int>(s.boxes.size()));
        for (HWND b : s.boxes)
        {
            const int id = GetDlgCtrlID(b);
            const int h = RowHeight(pane, id);
            const int x = IsGroupId(id) ? Px(pane, 2) : Px(pane, 22);
            if (dwp) dwp = DeferWindowPos(dwp, b, nullptr, x, y - s.scroll, rc.right - x, h,
                                          SWP_NOZORDER | SWP_NOACTIVATE);
            y += h;
        }
        if (dwp) EndDeferWindowPos(dwp);
        s.content = y;
        const int page = rc.bottom;
        const int maxScroll = (s.content > page) ? s.content - page : 0;
        if (s.scroll > maxScroll) { s.scroll = maxScroll; Layout(pane); return; }
        SCROLLINFO si{ sizeof(si), SIF_RANGE | SIF_PAGE | SIF_POS, 0, s.content > 0 ? s.content - 1 : 0,
                       static_cast<UINT>(page), s.scroll, 0 };
        SetScrollInfo(pane, SB_VERT, &si, TRUE);
    }

    // The rows move as one picture, so only those it uncovers are drawn.
    void ScrollTo(HWND pane, int pos)
    {
        State& s = *Of(pane);
        RECT rc; GetClientRect(pane, &rc);
        const int maxScroll = (s.content > rc.bottom) ? s.content - rc.bottom : 0;
        pos = pos < 0 ? 0 : (pos > maxScroll ? maxScroll : pos);
        if (pos == s.scroll) return;
        const int dy = s.scroll - pos;
        s.scroll = pos;
        ScrollWindowEx(pane, 0, dy, nullptr, nullptr, nullptr, nullptr, SW_SCROLLCHILDREN | SW_INVALIDATE | SW_ERASE);
        SCROLLINFO si{ sizeof(si), SIF_POS };
        si.nPos = pos;
        SetScrollInfo(pane, SB_VERT, &si, TRUE);
        UpdateWindow(pane);
    }

    // Keyboard focus on a box scrolled out of sight brings it into view.
    void EnsureVisible(HWND pane, HWND box)
    {
        RECT rc, br; GetClientRect(pane, &rc);
        GetWindowRect(box, &br);
        MapWindowPoints(nullptr, pane, reinterpret_cast<POINT*>(&br), 2);
        State& s = *Of(pane);
        if (br.top < 0)               ScrollTo(pane, s.scroll + br.top);
        else if (br.bottom > rc.bottom) ScrollTo(pane, s.scroll + br.bottom - rc.bottom);
    }

    void DrawBox(HWND pane, const DRAWITEMSTRUCT* di)
    {
        HWND h = di->hwndItem;
        const int id = static_cast<int>(di->CtlID);
        const int cs = static_cast<int>(GetWindowLongPtrW(h, GWLP_USERDATA));
        const bool disabled = (di->itemState & ODS_DISABLED) != 0;
        const bool focused  = (di->itemState & ODS_FOCUS) != 0;
        const bool hot = IsHot(h) && !disabled;
        const bool heading = IsGroupId(id);
        const State& s = *Of(pane);
        Buffered(di->hDC, di->rcItem, [&](HDC dc, const RECT& rc) {
            FillRect(dc, &rc, PageBrush());
            const int side = Px(h, 13);
            const int top = heading ? rc.bottom - Px(h, 6) - side : (rc.bottom - side) / 2;
            const RECT b{ 0, top, side, top + side };
            PaintCheckBox(h, dc, b, cs, hot, disabled);

            RECT t{ b.right + Px(h, 6), b.top - Px(h, 3), rc.right, b.bottom + Px(h, 3) };
            wchar_t name[96] = {};
            GetWindowTextW(h, name, 96);
            Text(h, dc, t, name, heading ? Face::Bold : Face::Body, disabled ? kDisabled : kText, ui::text::kVCentre);
            if (!heading)
            {
                const int k = id - kFirstEventId;
                const ev::Event& e = ev::At(k);
                const bool absent = !((s.available >> k) & 1);
                wchar_t note[128];
                _snwprintf_s(note, _TRUNCATE, L"%S%s", absent ? "not in this Excel" : e.shownAs,
                             (e.often && !absent) ? L"  \u26A1" : L"");
                RECT n = t;
                n.left = b.right + Px(h, 250);
                Text(h, dc, n, note, Face::Body, kDisabled, ui::text::kVCentre | ui::text::kEllipsis);
            }
            if (focused)
            {
                SIZE sz{};
                ui::text::Measure(name, heading ? Face::Bold : Face::Body, 0, DpiOf(h), sz);
                RECT f{ t.left - Px(h, 3), (t.top + t.bottom - sz.cy) / 2 - Px(h, 2),
                        t.left + sz.cx + Px(h, 3), (t.top + t.bottom + sz.cy) / 2 + Px(h, 2) };
                HPEN pen = CreatePen(PS_INSIDEFRAME, Line(h, 2), kHeavy);
                HGDIOBJ op = SelectObject(dc, pen);
                HGDIOBJ ob = SelectObject(dc, GetStockObject(NULL_BRUSH));
                Rectangle(dc, f.left, f.top, f.right, f.bottom);
                SelectObject(dc, ob); SelectObject(dc, op);
                DeleteObject(pen);
            } });
    }

    void Clicked(HWND pane, int id)
    {
        State& s = *Of(pane);
        if (!s.enabled) return;
        if (IsGroupId(id)) CycleGroup(s, id - kFirstGroupId);
        else if (IsEventId(id))
        {
            const int k = id - kFirstEventId;
            if (!((s.available >> k) & 1)) return;
            s.selected ^= ev::Mask(1) << k;
            // A single tick is a new mix: the heading's remembered one no longer means anything.
            s.remembers[static_cast<int>(ev::At(k).group)] = false;
        }
        else return;
        Refresh(pane);
        HWND dlg = GetParent(pane);
        SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(GetDlgCtrlID(pane), kChanged), reinterpret_cast<LPARAM>(pane));
        if (s.tip) SendMessageW(s.tip, TTM_UPDATE, 0, 0);
    }

    LRESULT CALLBACK PaneProc(HWND pane, UINT m, WPARAM wp, LPARAM lp)
    {
        State* s = Of(pane);
        switch (m)
        {
        case WM_SIZE:
            if (s) Layout(pane);
            return 0;
        case WM_DRAWITEM:
            DrawBox(pane, reinterpret_cast<const DRAWITEMSTRUCT*>(lp));
            return TRUE;
        case WM_COMMAND:
            // An owner-drawn button reports a quick second click as a double-click instead.
            if (HIWORD(wp) == BN_CLICKED || HIWORD(wp) == BN_DOUBLECLICKED) { Clicked(pane, LOWORD(wp)); return 0; }
            if (HIWORD(wp) == BN_SETFOCUS) { EnsureVisible(pane, reinterpret_cast<HWND>(lp)); return 0; }
            break;
        case WM_NOTIFY:
        {
            NMHDR* nh = reinterpret_cast<NMHDR*>(lp);
            if (nh && nh->code == TTN_GETDISPINFOW)
            {
                NMTTDISPINFOW* ti = reinterpret_cast<NMTTDISPINFOW*>(lp);
                static std::wstring text;
                text = TipOf(pane, GetDlgCtrlID(reinterpret_cast<HWND>(ti->hdr.idFrom)));
                ti->lpszText = const_cast<wchar_t*>(text.c_str());
                return 0;
            }
            break;
        }
        case WM_VSCROLL:
        {
            if (!s) break;
            SCROLLINFO si{ sizeof(si), SIF_ALL };
            GetScrollInfo(pane, SB_VERT, &si);
            const int line = Px(pane, 22);
            int pos = s->scroll;
            switch (LOWORD(wp))
            {
            case SB_LINEUP:        pos -= line; break;
            case SB_LINEDOWN:      pos += line; break;
            case SB_PAGEUP:        pos -= static_cast<int>(si.nPage); break;
            case SB_PAGEDOWN:      pos += static_cast<int>(si.nPage); break;
            case SB_THUMBTRACK:
            case SB_THUMBPOSITION: pos = si.nTrackPos; break;
            case SB_TOP:           pos = 0; break;
            case SB_BOTTOM:        pos = s->content; break;
            }
            ScrollTo(pane, pos);
            return 0;
        }
        case WM_MOUSEWHEEL:
            if (s) ScrollTo(pane, s->scroll - GET_WHEEL_DELTA_WPARAM(wp) * Px(pane, 22) * 3 / WHEEL_DELTA);
            return 0;
        case WM_ERASEBKGND:
        {
            RECT rc; GetClientRect(pane, &rc);
            FillRect(reinterpret_cast<HDC>(wp), &rc, PageBrush());
            return 1;
        }
        case WM_NCDESTROY:
            if (s)
            {
                if (s->acc)
                {
                    const MSAAPROPID props[] = { PROPID_ACC_ROLE, PROPID_ACC_STATE };
                    for (HWND b : s->boxes)
                        s->acc->ClearHwndProps(b, static_cast<DWORD>(OBJID_CLIENT), CHILDID_SELF, props, 2);
                    s->acc->Release();
                }
                if (s->tip) DestroyWindow(s->tip);
                delete s;
                SetWindowLongPtrW(pane, GWLP_USERDATA, 0);
            }
            break;
        }
        return DefWindowProcW(pane, m, wp, lp);
    }

    void Register()
    {
        static bool done = false;
        if (done) return;
        WNDCLASSW wc{};
        wc.lpfnWndProc = PaneProc;
        wc.hInstance = Self();
        wc.lpszClassName = kClass;
        wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
        done = RegisterClassW(&wc) != 0 || GetLastError() == ERROR_CLASS_ALREADY_EXISTS;
    }
}

HWND Create(HWND dlg, int id)
{
    Register();
    HWND pane = CreateWindowExW(WS_EX_CONTROLPARENT, kClass, L"", WS_CHILD | WS_VSCROLL | WS_CLIPCHILDREN,
                                0, 0, 10, 10, dlg, reinterpret_cast<HMENU>(static_cast<INT_PTR>(id)), Self(), nullptr);
    if (!pane) return nullptr;
    State* s = new State();
    SetWindowLongPtrW(pane, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(s));

    // Screen readers: absent when COM is not initialised on this thread, which only a test can be.
    CoCreateInstance(CLSID_AccPropServices, nullptr, CLSCTX_INPROC_SERVER, IID_IAccPropServices,
                     reinterpret_cast<void**>(&s->acc));

    INITCOMMONCONTROLSEX icc{ sizeof(icc), ICC_WIN95_CLASSES };
    InitCommonControlsEx(&icc);
    s->tip = CreateWindowExW(WS_EX_TOPMOST, TOOLTIPS_CLASSW, nullptr, WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
                             CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, pane, nullptr, Self(), nullptr);
    if (s->tip) SendMessageW(s->tip, TTM_SETMAXTIPWIDTH, 0, 400);

    const auto add = [&](int boxId, const wchar_t* text) {
        HWND b = CreateWindowExW(0, L"Button", text, WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_OWNERDRAW | BS_NOTIFY,
                                 0, 0, 10, 10, pane, reinterpret_cast<HMENU>(static_cast<INT_PTR>(boxId)), Self(), nullptr);
        if (!b) return;
        Subclass(b, ButtonProc);            // hover, as the dialog's other boxes have
        s->boxes.push_back(b);
        if (s->tip)
        {
            TOOLINFOW ti{ sizeof(ti) };
            ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
            ti.hwnd = pane;
            ti.uId = reinterpret_cast<UINT_PTR>(b);
            ti.lpszText = LPSTR_TEXTCALLBACKW;
            SendMessageW(s->tip, TTM_ADDTOOLW, 0, reinterpret_cast<LPARAM>(&ti));
        }
    };
    for (int g = 0; g < kGroups; ++g)
    {
        wchar_t name[64];
        _snwprintf_s(name, _TRUNCATE, L"%S", ev::GroupName(static_cast<ev::Group>(g)));
        add(kFirstGroupId + g, name);
        for (int k = 0; k < ev::Count(); ++k)
        {
            if (static_cast<int>(ev::At(k).group) != g) continue;
            wchar_t ename[64];
            _snwprintf_s(ename, _TRUNCATE, L"%S", ev::At(k).name);
            add(kFirstEventId + k, ename);
        }
    }
    return pane;
}

void Set(HWND pane, ev::Mask selected, ev::Mask available)
{
    State* s = Of(pane);
    if (!s) return;
    s->selected = selected;
    s->available = available;
    for (bool& r : s->remembers) r = false;    // a preset or a new session forgets every group's mix
    Refresh(pane);
}

ev::Mask Selected(HWND pane)
{
    const State* s = Of(pane);
    return s ? s->selected : 0;
}

void Enable(HWND pane, bool on)
{
    State* s = Of(pane);
    if (!s) return;
    s->enabled = on;
    Refresh(pane);
}

int GroupState(HWND pane, int group)
{
    const State* s = Of(pane);
    return s ? StateOf(*s, group) : 0;
}

void FillPresets(HWND combo)
{
    for (ev::Preset p : kPresets)
    {
        wchar_t w[32];
        core::WidenUtf8(ev::PresetName(p), w, 32);
        SendMessageW(combo, CB_ADDSTRING, 0, reinterpret_cast<LPARAM>(w));
    }
}

void ShowSummary(HWND pane, HWND combo, HWND count)
{
    const State* s = Of(pane);
    if (!s) return;
    const ev::Preset p = ev::PresetOf(s->selected, s->available);
    const LRESULT shown = SendMessageW(combo, CB_GETCURSEL, 0, 0);
    for (int i = 0; i < kPresetCount; ++i)
        if (kPresets[i] == p && i != shown) SendMessageW(combo, CB_SETCURSEL, i, 0);
    wchar_t t[48], was[48] = {};
    _snwprintf_s(t, _TRUNCATE, L"%d of %d events", Count(s->selected & s->available), Count(s->available));
    GetWindowTextW(count, was, 48);
    if (std::wcscmp(t, was) != 0) SetWindowTextW(count, t);
}

void ChoosePreset(HWND pane, HWND combo, HWND count)
{
    const State* s = Of(pane);
    if (!s) return;
    const int i = static_cast<int>(SendMessageW(combo, CB_GETCURSEL, 0, 0));
    if (i >= 0 && i < kPresetCount && kPresets[i] != ev::Preset::Custom)
        Set(pane, (ev::PresetMask(kPresets[i]) & s->available) | (s->selected & ~s->available), s->available);
    ShowSummary(pane, combo, count);
}

std::wstring TipOf(HWND pane, int id)
{
    const State* s = Of(pane);
    if (!s) return L"";
    if (IsGroupId(id))
    {
        const int g = id - kFirstGroupId;
        const int n = Count(GroupMask(g) & s->available);
        wchar_t t[64];
        switch (StateOf(*s, g))
        {
        case 1:  return L"Click to clear all";
        case 2:  _snwprintf_s(t, _TRUNCATE, L"Click to tick all %d", n); return t;
        default:
            if (s->remembers[g]) return L"Click to restore your selection";
            _snwprintf_s(t, _TRUNCATE, L"Click to tick all %d", n);
            return t;
        }
    }
    if (IsEventId(id))
    {
        wchar_t p[256];
        if (app::appevents::ParamsText(id - kFirstEventId, p, 256)) return std::wstring(L"Records: ") + p;
        wchar_t t[160];
        _snwprintf_s(t, _TRUNCATE, L"%S", ev::At(id - kFirstEventId).shownAs);
        return t;
    }
    return L"";
}
}
}
