// The Diagnostics dialog, opened for real with no Excel: the shipping dialog code and
// resources, a second thread driving it through its controls, and stubs for what the add-in's
// session would supply. The tables are read from this test's own process, so what the dialog
// shows can be checked against what the process is.
//
// Export is not pressed: the button writes the file and then reveals it in Explorer, which a
// sweep should not leave open. Everything it writes goes through diag::AsCsv and
// diag::WriteBytes, which are driven here directly.
#include "ui/diagnosticsdlg.h"
#include "ui/diagres.h"
#include "ui/gridlist.h"
#include "diag/snapshot.h"
#include "diag/tableout.h"
#include "core/log.h"

#include <windows.h>
#include <cstdio>
#include <cwchar>
#include <string>
#include <vector>

// ---- what the add-in's session would supply -------------------------------------------------

static bool g_armed = false;

namespace app
{
    bool IsArmed() { return g_armed; }
    const char* VersionText() { return "9.9.9"; }
}
namespace emit { namespace csv { std::wstring Path() { return L""; } } }
namespace core
{
    std::wstring EnsureAppSubdir(const wchar_t* leaf)
    {
        wchar_t temp[MAX_PATH] = {};
        GetTempPathW(MAX_PATH, temp);
        std::wstring dir = std::wstring(temp) + L"XRayXL_diagtest\\" + leaf;
        CreateDirectoryW((std::wstring(temp) + L"XRayXL_diagtest").c_str(), nullptr);
        CreateDirectoryW(dir.c_str(), nullptr);
        return dir;
    }
    namespace Log
    {
        static Level g_level = Level::Info;
        std::wstring Path() { return L"C:\\stub\\XRayXL.log"; }
        void  SetLevel(Level lvl) { g_level = lvl; }
        Level GetLevel() { return g_level; }
        bool  LevelFromText(const char* text, Level& out)
        {
            const char* const names[] = { "DEBUG", "INFO", "WARNING", "ERROR" };
            for (int i = 0; i < 4; ++i)
                if (_stricmp(text, names[i]) == 0) { out = static_cast<Level>(i); return true; }
            return false;
        }
        void Write(Level, const std::string&) {}
        void Debug(const std::string&) {}
        void Info(const std::string&) {}
        void Warning(const std::string& m) { std::printf("  (log) warning: %s\n", m.c_str()); }
        void Error(const std::string& m) { std::printf("  (log) error: %s\n", m.c_str()); }
        void Note(const std::string&) {}
    }
}

// ---- checks -------------------------------------------------------------------------------

static int g_fail = 0;

static void Check(bool ok, const char* what, const std::wstring& detail = L"")
{
    std::printf("[%s] %s", ok ? "PASS" : "FAIL", what);
    if (!ok && !detail.empty()) std::wprintf(L"  (%ls)", detail.substr(0, 200).c_str());
    std::printf("\n");
    if (!ok) ++g_fail;
}

static void Check(bool ok, const char* what, const std::string& detail)
{
    Check(ok, what, std::wstring(detail.begin(), detail.end()));
}

// ---- driving the dialog from a second thread ----------------------------------------------

static HWND FindDialog()
{
    for (int i = 0; i < 200; ++i)
    {
        for (HWND h = nullptr; (h = FindWindowExW(nullptr, h, L"#32770", L"XRayXL Diagnostics")) != nullptr; )
        {
            DWORD pid = 0;
            GetWindowThreadProcessId(h, &pid);
            if (pid == GetCurrentProcessId() && IsWindowVisible(h)) return h;
        }
        Sleep(50);
    }
    return nullptr;
}

static std::wstring TextOf(HWND dlg, int id)
{
    HWND c = GetDlgItem(dlg, id);
    if (!c) return L"";
    const LRESULT n = SendMessageW(c, WM_GETTEXTLENGTH, 0, 0);
    std::wstring s(static_cast<size_t>(n) + 1, L'\0');
    SendMessageW(c, WM_GETTEXT, static_cast<WPARAM>(s.size()), reinterpret_cast<LPARAM>(&s[0]));
    s.resize(static_cast<size_t>(n));
    return s;
}

static HWND List(HWND dlg) { return GetDlgItem(dlg, IDC_DIAG_LIST); }

static void ShowPage(HWND dlg, const wchar_t* name)
{
    HWND list = GetDlgItem(dlg, IDC_DIAG_CATEGORIES);
    const LRESULT at = SendMessageW(list, LB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(name));
    SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(at), 0);
    SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_DIAG_CATEGORIES, LBN_SELCHANGE), reinterpret_cast<LPARAM>(list));
}

static void TypeSearch(HWND dlg, const wchar_t* text)
{
    SetDlgItemTextW(dlg, IDC_DIAG_SEARCH, text);
    SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_DIAG_SEARCH, EN_CHANGE),
                 reinterpret_cast<LPARAM>(GetDlgItem(dlg, IDC_DIAG_SEARCH)));
}

// A header click at the middle of column `col`, as the mouse would deliver it.
static void ClickHeader(HWND dlg, int col)
{
    HWND list = List(dlg);
    const diag::Table t = ui::grid::Shown(list);
    int x = 0;
    for (int i = 0; i < col && i < static_cast<int>(t.columns.size()); ++i) x += t.columns[i].width96;
    x += (col < static_cast<int>(t.columns.size()) ? t.columns[col].width96 : 40) / 2;
    SendMessageW(list, WM_LBUTTONDOWN, 0, MAKELPARAM(x, 12));
    SendMessageW(list, WM_LBUTTONUP, 0, MAKELPARAM(x, 12));
}

// What the window reports the cursor is over, at a point given in client coordinates.
static LRESULT HitTestAt(HWND dlg, int clientX, int clientY)
{
    POINT p{ clientX, clientY };
    ClientToScreen(dlg, &p);
    return SendMessageW(dlg, WM_NCHITTEST, 0, MAKELPARAM(p.x, p.y));
}

// A click on row `row`, as the mouse would deliver it: the modifiers ride in wParam.
static void ClickRow(HWND dlg, int row, bool shift, bool ctrl)
{
    const WPARAM mods = (shift ? MK_SHIFT : 0) | (ctrl ? MK_CONTROL : 0);
    const int y = 25 + row * 17 + 8;              // header, rows, mid-row
    SendMessageW(List(dlg), WM_LBUTTONDOWN, mods, MAKELPARAM(30, y));
    SendMessageW(List(dlg), WM_LBUTTONUP, mods, MAKELPARAM(30, y));
}

static void Press(HWND dlg, int id) { SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(id, BN_CLICKED), 0); }

template <class Drive>
static bool Session(const char* name, Drive drive)
{
    std::printf("-- %s\n", name);
    struct Ctx { Drive* d; const char* name; } ctx{ &drive, name };
    HANDLE t = CreateThread(nullptr, 0, [](LPVOID p) -> DWORD {
        Ctx* c = static_cast<Ctx*>(p);
        HWND dlg = FindDialog();
        const std::string opened = std::string("the dialog opened: ") + c->name;
        Check(dlg != nullptr, opened.c_str());
        if (dlg) (*c->d)(dlg);
        return 0; }, &ctx, 0, nullptr);
    const bool ok = ui::diagnostics::Show(nullptr);
    WaitForSingleObject(t, 30000);
    CloseHandle(t);
    return ok;
}

static bool Contains(const std::wstring& s, const wchar_t* needle) { return s.find(needle) != std::wstring::npos; }
static bool Contains(const std::string& s, const char* needle) { return s.find(needle) != std::string::npos; }

static std::wstring Lower(std::wstring s)
{
    for (wchar_t& c : s) c = static_cast<wchar_t>(towlower(c));
    return s;
}

// ---- what the dialog must leave behind: nothing --------------------------------------------

// The list is our own window class. If it outlived the module, an XLL Excel unloaded and loaded
// again would meet a class whose window procedure is no longer mapped.
static bool GridClassRegistered()
{
    WNDCLASSEXW wc{};
    wc.cbSize = sizeof(wc);
    return GetClassInfoExW(GetModuleHandleW(nullptr), L"XRayXLGridList", &wc) != FALSE;
}

struct Resources
{
    DWORD gdi = 0, user = 0, handles = 0;
};

static Resources Used()
{
    Resources r;
    r.gdi  = GetGuiResources(GetCurrentProcess(), GR_GDIOBJECTS);
    r.user = GetGuiResources(GetCurrentProcess(), GR_USEROBJECTS);
    GetProcessHandleCount(GetCurrentProcess(), &r.handles);
    return r;
}

// ---- the tables on their own, with no dialog ------------------------------------------------

static void TableCases()
{
    std::printf("-- the tables\n");

    const diag::Table modules = diag::Modules();
    Check(modules.columns.size() == 5, "Modules has five columns");
    Check(modules.columns[3].rightAlign && !modules.columns[0].rightAlign,
          "Size reads against its own edge; the rest read from the left");
    Check(modules.columns[0].title == L"Module" && modules.columns[1].title == L"Version" &&
          modules.columns[2].title == L"Date Modified" && modules.columns[3].title == L"Size" &&
          modules.columns[4].title == L"Path",
          "the module columns are Module, Version, Date Modified, Size, Path");
    Check(modules.rows.size() > 5, "more than five modules are loaded into this process");

    wchar_t self[MAX_PATH] = {};
    GetModuleFileNameW(nullptr, self, MAX_PATH);
    Check(!modules.rows.empty() && Lower(modules.rows[0][4]) == Lower(self),
          "the first row is the process image itself", modules.rows.empty() ? L"" : modules.rows[0][4]);

    bool kernel = false, sorted = true;
    for (size_t i = 1; i < modules.rows.size(); ++i)
    {
        if (Lower(modules.rows[i][0]) == L"kernel32.dll") kernel = true;
        if (i > 1 && _wcsicmp(modules.rows[i - 1][0].c_str(), modules.rows[i][0].c_str()) > 0) sorted = false;
    }
    Check(kernel, "KERNEL32.DLL is in the list");
    Check(sorted, "the modules after the process image are in name order");

    bool versioned = false, dated = false;
    for (const std::vector<std::wstring>& r : modules.rows)
    {
        if (Lower(r[0]) != L"kernel32.dll") continue;
        versioned = !r[1].empty();
        dated = r[2].size() == 19 && r[2][4] == L'-' && r[2][10] == L' ';
        Check(Contains(Lower(r[4]), L"kernel32.dll"), "its path names the file", r[4]);
    }
    Check(versioned, "KERNEL32.DLL carries a version");
    bool tagged = false;
    for (const std::vector<std::wstring>& r : modules.rows)
        if (Contains(r[1], L"(")) tagged = true;
    Check(!tagged, "no version carries Microsoft's build-lab tag");
    Check(dated, "a date modified reads yyyy-mm-dd hh:mm:ss");

    SetEnvironmentVariableW(L"XRAYXL_DIAGTEST_MARKER", L"a value with = and spaces");
    const diag::Table env = diag::Environment();
    bool found = false;
    for (const std::vector<std::wstring>& r : env.rows)
        if (r[0] == L"XRAYXL_DIAGTEST_MARKER") { found = r[1] == L"a value with = and spaces"; }
    Check(env.columns.size() == 2 && env.columns[0].title == L"Variable", "Environment is Variable and Value");
    Check(found, "a variable set before the read appears with its whole value");

    diag::Session session;
    session.version = "9.9.9";
    session.armed = false;
    session.logFile = L"C:\\stub\\XRayXL.log";
    const diag::Table proc = diag::Process(session);
    wchar_t pid[16];
    _snwprintf_s(pid, _TRUNCATE, L"%lu", GetCurrentProcessId());
    bool sawPid = false, sawHandles = false, sawGdi = false, sawWorking = false, sawVersion = false;
    for (const std::vector<std::wstring>& r : proc.rows)
    {
        if (r[0] == L"Process ID")     sawPid = (r[1] == pid);
        if (r[0] == L"Handles")        sawHandles = !r[1].empty();
        if (r[0] == L"GDI objects")    sawGdi = !r[1].empty();
        if (r[0] == L"Working set")    sawWorking = Contains(r[1], L"KB");
        if (r[0] == L"XRayXL version") sawVersion = (r[1] == L"9.9.9");
    }
    Check(sawPid, "Process reports this process's id");
    Check(sawHandles, "Process reports a handle count");
    Check(sawGdi, "Process reports a GDI object count");
    Check(sawWorking, "Process reports a working set in KB");
    Check(sawVersion, "Process reports the XRayXL version it was given");
}

// ---- what Export would write ----------------------------------------------------------------

static void ExportCases()
{
    std::printf("-- export\n");

    diag::Table t;
    t.columns = { { L"Name", 100, false }, { L"Note", 100, false } };
    t.rows.push_back({ L"plain", L"nothing special" });
    t.rows.push_back({ L"comma,and \"quote\"", L"line\r\nbreak" });

    const std::string csv = diag::AsCsv(t);
    Check(csv.rfind("\xEF\xBB\xBF", 0) == 0, "the CSV starts with a BOM, so Excel reads it as UTF-8");
    Check(Contains(csv, "Name,Note\r\n"), "the CSV header is the column titles");
    Check(Contains(csv, "\"comma,and \"\"quote\"\"\""), "a comma and a quote are quoted and doubled", csv);
    Check(Contains(csv, "\"line\r\nbreak\""), "a line break is quoted rather than ending the row");

    const std::wstring tabbed = diag::AsTabbed(t);
    Check(Contains(tabbed, L"Name\tNote\r\n"), "the clipboard text is tab separated");

    const std::wstring path = diag::ExportPath(L"Modules");
    Check(!path.empty() && Contains(path, L"XRayXL_Modules_") && Contains(path, L".csv"),
          "the export path names the page", path);
    Check(diag::WriteBytes(path, csv), "the export file is written", path);

    HANDLE h = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    DWORD size = (h != INVALID_HANDLE_VALUE) ? GetFileSize(h, nullptr) : 0;
    if (h != INVALID_HANDLE_VALUE) CloseHandle(h);
    Check(size == csv.size(), "the file on disk is the bytes that were written");
    DeleteFileW(path.c_str());
}

// Opening and closing the dialog many times must cost nothing that is not given back. A GUI
// component that leaks does it per cycle, so a climb over twenty is the signal; a handful of
// one-off allocations on the first open is not, which is what the warm-up absorbs.
static void TeardownCases()
{
    std::printf("-- teardown\n");

    Check(!GridClassRegistered(), "the list's window class is not registered before the first open");

    bool seenWhileOpen = false;
    Session("teardown: the class exists while the dialog is up", [&](HWND dlg)
    {
        seenWhileOpen = GridClassRegistered();
        Press(dlg, IDCANCEL);
    });
    Check(seenWhileOpen, "the class is registered while the dialog is open");
    Check(!GridClassRegistered(), "and is gone once it has closed");

    for (int warm = 0; warm < 3; ++warm)
        Session("teardown: warm-up", [](HWND dlg) { Press(dlg, IDCANCEL); });

    const Resources before = Used();
    constexpr int kCycles = 20;
    int classLeft = 0;
    for (int i = 0; i < kCycles; ++i)
    {
        Session("teardown: cycle", [](HWND dlg) { Press(dlg, IDCANCEL); });
        if (GridClassRegistered()) ++classLeft;
    }
    const Resources after = Used();

    Check(classLeft == 0, "the window class is unregistered after every one of twenty opens");

    wchar_t detail[160];
    _snwprintf_s(detail, _TRUNCATE, L"gdi %lu->%lu, user %lu->%lu, handles %lu->%lu over %d opens",
                 before.gdi, after.gdi, before.user, after.user,
                 before.handles, after.handles, kCycles);
    // One per cycle would show as twenty; a few from a cache settling would not.
    Check(after.gdi <= before.gdi + 4, "GDI objects do not climb with each open", detail);
    Check(after.user <= before.user + 4, "USER objects do not climb with each open", detail);
    Check(after.handles <= before.handles + 8, "handles do not climb with each open", detail);
    std::wprintf(L"       %ls\n", detail);
}

int main()
{
    TableCases();
    ExportCases();
    TeardownCases();

    Session("the three pages, search, sort and the detail line", [](HWND dlg)
    {
        HWND cats = GetDlgItem(dlg, IDC_DIAG_CATEGORIES);
        std::wstring names;
        const LRESULT n = SendMessageW(cats, LB_GETCOUNT, 0, 0);
        for (LRESULT i = 0; i < n; ++i)
        {
            wchar_t t[64] = {};
            SendMessageW(cats, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(t));
            names += (i ? L"," : L"") + std::wstring(t);
        }
        Check(names == L"Modules,Environment,Process", "the categories are Modules, Environment, Process", names);
        Check(List(dlg) != nullptr, "the list control was created");

        // ---- Modules, and the search that the issue asks for ----
        const int all = ui::grid::RowCount(List(dlg));
        Check(all > 5, "Modules opens showing every loaded module");

        TypeSearch(dlg, L"kernel32");
        const int narrowed = ui::grid::RowCount(List(dlg));
        Check(narrowed >= 1 && narrowed < all, "searching narrows the list");
        const diag::Table hits = ui::grid::Shown(List(dlg));
        bool allMatch = !hits.rows.empty();
        for (const std::vector<std::wstring>& r : hits.rows)
        {
            bool any = false;
            for (const std::wstring& cell : r) any = any || Contains(Lower(cell), L"kernel32");
            allMatch = allMatch && any;
        }
        Check(allMatch, "every row left holds the search text in some column");

        TypeSearch(dlg, L"KERNEL32");
        Check(ui::grid::RowCount(List(dlg)) == narrowed, "the search ignores case");

        TypeSearch(dlg, L"zzz-no-such-module");
        Check(ui::grid::RowCount(List(dlg)) == 0, "a search that matches nothing empties the list");

        TypeSearch(dlg, L"");
        Check(ui::grid::RowCount(List(dlg)) == all, "clearing the search brings every row back");

        // ---- sorting, by the header click Excel's own list takes ----
        ClickHeader(dlg, 0);
        const diag::Table up = ui::grid::Shown(List(dlg));
        bool ascending = true;
        for (size_t i = 1; i < up.rows.size(); ++i)
            if (_wcsicmp(up.rows[i - 1][0].c_str(), up.rows[i][0].c_str()) > 0) ascending = false;
        Check(ascending && up.rows.size() > 1, "clicking Module sorts up");

        ClickHeader(dlg, 0);
        const diag::Table down = ui::grid::Shown(List(dlg));
        bool descending = true;
        for (size_t i = 1; i < down.rows.size(); ++i)
            if (_wcsicmp(down.rows[i - 1][0].c_str(), down.rows[i][0].c_str()) < 0) descending = false;
        Check(descending, "clicking it again sorts down");

        // Size is grouped digits, and must sort as a number rather than as text.
        ClickHeader(dlg, 3);
        const diag::Table bySize = ui::grid::Shown(List(dlg));
        bool numeric = true;
        double last = -1;
        for (const std::vector<std::wstring>& r : bySize.rows)
        {
            std::wstring plain;
            for (wchar_t c : r[3]) if (c != L',') plain += c;
            if (plain.empty()) continue;
            const double v = _wtof(plain.c_str());
            if (v + 0.5 < last) numeric = false;
            last = v;
        }
        Check(numeric, "Size sorts by size, not by leading digit");

        // ---- choosing rows: one, a Ctrl-added one, a Shift range, and all of them ----
        SendMessageW(List(dlg), WM_KEYDOWN, VK_HOME, 0);
        Check(ui::grid::Focused(List(dlg)) == 0, "Home moves the focus to the first row");
        Check(ui::grid::ChosenCount(List(dlg)) == 1, "and chooses it");
        const diag::Table one = ui::grid::Chosen(List(dlg));
        Check(one.rows.size() == 1 && one.rows[0][0] == ui::grid::Shown(List(dlg)).rows[0][0],
              "the chosen row is the focused one");

        SendMessageW(List(dlg), WM_KEYDOWN, VK_DOWN, 0);
        Check(ui::grid::Focused(List(dlg)) == 1 && ui::grid::ChosenCount(List(dlg)) == 1,
              "Down moves the choice rather than adding to it");

        ClickRow(dlg, 3, false, false);
        Check(ui::grid::ChosenCount(List(dlg)) == 1, "a plain click chooses one row");
        ClickRow(dlg, 5, false, true);
        Check(ui::grid::ChosenCount(List(dlg)) == 2, "Ctrl+click adds a row");
        ClickRow(dlg, 5, false, true);
        Check(ui::grid::ChosenCount(List(dlg)) == 1, "Ctrl+click again takes it away");
        ClickRow(dlg, 2, false, false);
        ClickRow(dlg, 6, true, false);
        Check(ui::grid::ChosenCount(List(dlg)) == 5, "Shift+click takes the range from the anchor");
        const diag::Table range = ui::grid::Chosen(List(dlg));
        Check(range.rows.size() == 5 && range.rows[0][0] == ui::grid::Shown(List(dlg)).rows[2][0] &&
              range.rows[4][0] == ui::grid::Shown(List(dlg)).rows[6][0],
              "the range is the rows between them, in view order");

        SendMessageW(List(dlg), WM_KEYDOWN, 'A', 0);          // without Ctrl held: not Select All
        Check(ui::grid::ChosenCount(List(dlg)) == 5, "a bare A does not select everything");
        ui::grid::SelectAll(List(dlg));
        Check(ui::grid::ChosenCount(List(dlg)) == ui::grid::RowCount(List(dlg)),
              "Select All takes every row shown");

        // a search narrows what Select All can reach
        TypeSearch(dlg, L"kernel32");
        ui::grid::SelectAll(List(dlg));
        Check(ui::grid::ChosenCount(List(dlg)) == narrowed,
              "Select All after a search takes only what is shown");
        Check(static_cast<int>(ui::grid::Chosen(List(dlg)).rows.size()) == narrowed,
              "and that is what Copy and Export would be given");
        TypeSearch(dlg, L"");

        // ---- Environment ----
        ShowPage(dlg, L"Environment");
        Check(ui::grid::Shown(List(dlg)).columns.size() == 2, "Environment shows two columns");
        TypeSearch(dlg, L"XRAYXL_DIAGTEST_MARKER");
        const diag::Table marker = ui::grid::Shown(List(dlg));
        Check(marker.rows.size() == 1 && marker.rows[0][1] == L"a value with = and spaces",
              "Environment carries the value the process had when the dialog opened",
              marker.rows.empty() ? L"" : marker.rows[0][1]);

        // ---- Process ----
        ShowPage(dlg, L"Process");
        Check(TextOf(dlg, IDC_DIAG_SEARCH).empty(), "changing page clears the search");
        const diag::Table proc = ui::grid::Shown(List(dlg));
        Check(proc.rows.size() > 10, "Process shows its metrics");
        Check(ui::grid::RowCount(List(dlg)) == static_cast<int>(proc.rows.size()),
              "the page opens unfiltered");

        // ---- the three pages are one reading, taken when the dialog opened ----
        SetEnvironmentVariableW(L"XRAYXL_DIAGTEST_LATE", L"added after the dialog opened");
        ShowPage(dlg, L"Environment");
        TypeSearch(dlg, L"XRAYXL_DIAGTEST_LATE");
        Check(ui::grid::RowCount(List(dlg)) == 0,
              "a variable set after the dialog opened does not appear: the pages are static");
        TypeSearch(dlg, L"XRAYXL_DIAGTEST_MARKER");
        Check(ui::grid::RowCount(List(dlg)) == 1, "what was there when it opened is still there");
        Check(ui::grid::ChosenCount(List(dlg)) == 0, "returning to a page chooses nothing");

        // ---- the resize grip's square reports itself as the sizing corner ----
        RECT client{};
        GetClientRect(dlg, &client);
        // the grip is 14 logical pixels square, flush into the bottom-right of the client
        Check(HitTestAt(dlg, client.right - 4, client.bottom - 4) == HTBOTTOMRIGHT,
              "the corner of the grip is the sizing corner");
        Check(HitTestAt(dlg, client.right - 12, client.bottom - 12) == HTBOTTOMRIGHT,
              "so is the far side of the dots, not just the border beyond them");
        Check(HitTestAt(dlg, client.right - 40, client.bottom - 40) != HTBOTTOMRIGHT,
              "the client away from the grip is not");
        Check(HitTestAt(dlg, client.right / 2, client.bottom / 2) != HTBOTTOMRIGHT,
              "nor is the middle of the dialog");

        Press(dlg, IDCANCEL);
    });

    std::printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
