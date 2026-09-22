// The Options dialog, opened for real with no Excel: the shipping dialog code and resources, a
// second thread driving it through its controls, and stubs for what the add-in's session would
// supply. Plus the text reflow the About and Notices pages use, on its own.
//
// Excel is not needed and not wanted: the harness's dialog watchdog would click away a dialog
// shown inside a test's Excel before it could be driven.
#include "ui/optionsdlg.h"
#include "ui/optionsres.h"
#include "ui/reflow.h"
#include "ui/dlgexcelstyle.h"
#include "app/settings.h"
#include "core/tracemodes.h"
#include "core/log.h"

#include <windows.h>
#include <cstdio>
#include <cwchar>
#include <string>
#include <vector>

// ---- what the add-in's session would supply -------------------------------------------------

static bool g_armed = false;
static std::vector<std::string> g_notes;     // what the dialog logged

namespace app
{
    bool IsArmed() { return g_armed; }
    const char* VersionText() { return "9.9.9"; }
}
namespace emit { namespace csv { std::wstring Path() { return L""; } } }
namespace core
{
    std::wstring EnsureAppSubdir(const wchar_t* leaf) { return std::wstring(L"C:\\stub\\") + leaf; }
    namespace Log
    {
        static Level g_level = Level::Info;
        std::wstring Path() { return L"C:\\stub\\XRayXL.log"; }
        void  SetLevel(Level lvl) { g_level = lvl; }
        Level GetLevel() { return g_level; }
        const char* LevelName(Level lvl)
        {
            const char* const names[] = { "DEBUG", "INFO", "WARNING", "ERROR" };
            return names[static_cast<int>(lvl)];
        }
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
        void Note(const std::string& m) { g_notes.push_back(m); }
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

// ---- driving the dialog from a second thread ----------------------------------------------

static HWND FindDialog()
{
    for (int i = 0; i < 200; ++i)
    {
        for (HWND h = nullptr; (h = FindWindowExW(nullptr, h, L"#32770", L"XRayXL Options")) != nullptr; )
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

static bool Visible(HWND dlg, int id) { HWND c = GetDlgItem(dlg, id); return c && IsWindowVisible(c); }
static bool Enabled(HWND dlg, int id) { HWND c = GetDlgItem(dlg, id); return c && IsWindowEnabled(c); }
static bool Checked(HWND dlg, int id) { return ui::excelstyle::IsChecked(dlg, id); }

// The page the user would get by clicking its name in the list.
static void ShowPage(HWND dlg, const wchar_t* name)
{
    HWND list = GetDlgItem(dlg, IDC_CATEGORIES);
    const LRESULT at = SendMessageW(list, LB_FINDSTRINGEXACT, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(name));
    SendMessageW(list, LB_SETCURSEL, static_cast<WPARAM>(at), 0);
    SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(IDC_CATEGORIES, LBN_SELCHANGE), reinterpret_cast<LPARAM>(list));
}

static std::wstring ComboItem(HWND dlg, int id, int i)
{
    wchar_t t[64] = {};
    SendDlgItemMessageW(dlg, id, CB_GETLBTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(t));
    return t;
}

static void Press(HWND dlg, int id) { SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(id, BN_CLICKED), 0); }

// A drop-down choice the way a user makes one: the selection, then the notification.
static void Choose(HWND dlg, int id, int index)
{
    SendDlgItemMessageW(dlg, id, CB_SETCURSEL, static_cast<WPARAM>(index), 0);
    SendMessageW(dlg, WM_COMMAND, MAKEWPARAM(id, CBN_SELCHANGE), reinterpret_cast<LPARAM>(GetDlgItem(dlg, id)));
}

static std::wstring Wide(const std::string& s) { return std::wstring(s.begin(), s.end()); }

static bool Logged(const char* needle)
{
    for (const std::string& n : g_notes) if (n.find(needle) != std::string::npos) return true;
    return false;
}

// One opening of the dialog: `drive` runs on a second thread against the dialog while this
// thread sits in its modal loop, and must close it. Returns what Show returned.
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
    const bool applied = ui::options::Show(nullptr);
    WaitForSingleObject(t, 30000);
    CloseHandle(t);
    return applied;
}

static bool Contains(const std::wstring& s, const wchar_t* needle) { return s.find(needle) != std::wstring::npos; }

static size_t LongestRun(const std::wstring& s, wchar_t c)
{
    size_t best = 0, run = 0;
    for (wchar_t x : s) { run = (x == c) ? run + 1 : 0; if (run > best) best = run; }
    return best;
}

static void SettingsCases()
{
    std::printf("-- settings list and changes\n");
    using app::settings::Snapshot;
    const Snapshot a = { { "XLL ARGS", "TRUE" }, { "FORMAT", "CSV" }, { "LOGLEVEL", "INFO" } };
    const Snapshot b = { { "XLL ARGS", "FALSE" }, { "FORMAT", "CSV" }, { "LOGLEVEL", "DEBUG" } };
    Check(app::settings::List(a) == "XLL ARGS=TRUE, FORMAT=CSV, LOGLEVEL=INFO", "a snapshot lists as NAME=VALUE",
          Wide(app::settings::List(a)));
    Check(app::settings::Changes(a, b) == "XLL ARGS TRUE -> FALSE; LOGLEVEL INFO -> DEBUG",
          "changes name only what differs, old to new", Wide(app::settings::Changes(a, b)));
    Check(app::settings::Changes(a, a).empty(), "no change is an empty string");

    const std::string all = app::settings::List(app::settings::Take());
    for (const char* name : { "XLL DEPTH=", "XLL ARGS=", "XLL RETVAL=", "VBA DEPTH=", "VBA ARGS=", "VBA RETVAL=",
                              "VBA OBJECTS=", "VBA BREAKPOINTS=", "BUFFERSIZE=", "BUFFERWHENFULL=", "FORMAT=", "LOGLEVEL=" })
        Check(all.find(name) != std::string::npos, (std::string("the live list has ") + name).c_str(), Wide(all));
}

static void ReflowCases()
{
    std::printf("-- reflow\n");
    using ui::text::Reflow;
    Check(Reflow("one\ntwo\n\nthree\n") == "one two\r\n\r\nthree", "hard-wrapped lines join; a blank line ends a paragraph");
    Check(Reflow("a\r\nb\r\n") == "a b", "CRLF input joins the same way");
    Check(Reflow("Title\n=====\ntext\n") == "Title\r\n========================\r\ntext",
          "a rule keeps its own line, is drawn 24 wide, and the next line starts fresh");
    Check(Reflow("met:\n\n 1. first\n    more\n 2. second\n") == "met:\r\n\r\n1. first more\r\n2. second",
          "a numbered clause starts its own line; its continuation joins it");
    Check(Reflow("see <https://x.org/>\nnext\n") == "see <https://x.org/>\r\nnext",
          "a line ending in an address keeps its break");
    Check(Reflow("") == "", "an empty text stays empty");
    Check(Reflow("2007 was\na year\n") == "2007 was a year", "a line starting with digits but no dot is not a clause");
    Check(Reflow("--x\n") == "--x", "dashes followed by text are not a rule");
}

int main()
{
    ReflowCases();
    SettingsCases();

    core::modes::SetFormat(core::modes::Format::Csv);

    // ---- disarmed: every page, the format moved to Output, Apply applies it ------------------
    const bool applied = Session("disarmed: pages, texts, format, Apply", [](HWND dlg)
    {
        HWND list = GetDlgItem(dlg, IDC_CATEGORIES);
        std::wstring names;
        const LRESULT n = SendMessageW(list, LB_GETCOUNT, 0, 0);
        for (LRESULT i = 0; i < n; ++i)
        {
            wchar_t t[64] = {};
            SendMessageW(list, LB_GETTEXT, static_cast<WPARAM>(i), reinterpret_cast<LPARAM>(t));
            names += (i ? L"," : L"") + std::wstring(t);
        }
        Check(names == L"Capture,Output,Advanced,About,Notices", "the categories are Capture, Output, Advanced, About, Notices", names);

        ShowPage(dlg, L"Output");
        Check(Visible(dlg, IDC_OUT_FMT) && Visible(dlg, IDC_OUT_FMTLBL), "Output shows the Format drop-down");
        Check(SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_GETCOUNT, 0, 0) == 2 &&
              ComboItem(dlg, IDC_OUT_FMT, 0) == L"CSV" && ComboItem(dlg, IDC_OUT_FMT, 1) == L"JSON Lines",
              "Format offers CSV and JSON Lines");
        Check(SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_GETCURSEL, 0, 0) == 0, "Format shows the current setting, CSV");
        Check(Enabled(dlg, IDC_OUT_FMT), "Format can be changed while disarmed");
        Check(TextOf(dlg, IDC_OUT_HDR) == L"Where the trace goes, and in what format.", "Output's title", TextOf(dlg, IDC_OUT_HDR));

        ShowPage(dlg, L"Advanced");
        Check(!Visible(dlg, IDC_OUT_FMT), "Advanced does not show the Format drop-down");
        Check(Visible(dlg, IDC_ADV_FULL) && Visible(dlg, IDC_ADV_LVL), "Advanced shows the buffer and the log level");
        Check(TextOf(dlg, IDC_ADV_SEC1) == L"Trace file", "Advanced's first section is Trace file", TextOf(dlg, IDC_ADV_SEC1));
        Check(Visible(dlg, IDC_ADV_BRK) && Contains(TextOf(dlg, IDC_ADV_BRK), L"breaks"),
              "Advanced offers the breaks column", TextOf(dlg, IDC_ADV_BRK));

        ShowPage(dlg, L"About");
        const std::wstring lic = TextOf(dlg, IDC_ABT_LICENSE);
        Check(Visible(dlg, IDC_ABT_LICENSE) && lic.rfind(L"GNU GENERAL PUBLIC LICENSE", 0) == 0,
              "About shows the embedded licence", lic.substr(0, 60));
        Check(Contains(lic, L"Everyone is permitted to copy and distribute verbatim copies of this license document, "
                            L"but changing it is not allowed."),
              "the licence's hard-wrapped lines are rejoined");
        Check(GetDlgItem(dlg, 1604) == nullptr, "About has no Third Party Notices link");
        Check(Contains(TextOf(dlg, IDC_ABT_HDR), L"9.9.9"), "About's title carries the version", TextOf(dlg, IDC_ABT_HDR));
        Check(!Visible(dlg, IDC_ARMEDNOTE), "About shows no armed note");

        ShowPage(dlg, L"Notices");
        const std::wstring notes = TextOf(dlg, IDC_NOT_TEXT);
        Check(Visible(dlg, IDC_NOT_HDR) && TextOf(dlg, IDC_NOT_HDR) == L"Third Party Notices", "Notices has its title",
              TextOf(dlg, IDC_NOT_HDR));
        Check(Visible(dlg, IDC_NOT_TEXT) && Contains(notes, L"MinHook") && Contains(notes, L"Tsuda Kageyu"),
              "Notices shows the embedded notices, MinHook's included", notes.substr(0, 80));
        Check(Contains(notes, L"\r\n1. Redistributions of source code") && Contains(notes, L"\r\n2. Redistributions in binary form"),
              "each numbered condition starts its own line");
        Check(LongestRun(notes, L'=') == 24 && LongestRun(notes, L'-') <= 24, "rules are drawn short enough not to wrap");
        Check(!Visible(dlg, IDC_ABT_LICENSE), "the licence box is not on the Notices page");
        Check(!Visible(dlg, IDC_ARMEDNOTE), "Notices shows no armed note");

        Check(TextOf(dlg, IDOK) == L"Apply", "the button is called Apply", TextOf(dlg, IDOK));
        Check(!Enabled(dlg, IDOK), "Apply is greyed while nothing differs");
        Press(dlg, IDOK);
        Check(IsWindowVisible(dlg) != FALSE, "pressing Apply with nothing to apply leaves the dialog open");

        ShowPage(dlg, L"Output");
        Choose(dlg, IDC_OUT_FMT, 1);
        Check(Enabled(dlg, IDOK), "choosing JSON Lines lights Apply");
        Choose(dlg, IDC_OUT_FMT, 0);
        Check(!Enabled(dlg, IDOK), "choosing CSV again greys it");

        ShowPage(dlg, L"Advanced");
        wchar_t buf[32] = {};
        GetDlgItemTextW(dlg, IDC_ADV_BUF, buf, 32);
        std::wstring same;                              // "64MB" as "64 mb"
        for (const wchar_t* p = buf; *p; ++p)
        {
            if (*p >= L'A' && *p <= L'Z' && (p == buf || !(p[-1] >= L'A' && p[-1] <= L'Z'))) same += L' ';
            same += static_cast<wchar_t>(towlower(*p));
        }
        SetDlgItemTextW(dlg, IDC_ADV_BUF, same.c_str());
        Check(!Enabled(dlg, IDOK), "the same buffer size written differently is no change", buf);
        SetDlgItemTextW(dlg, IDC_ADV_BUF, L"lots");
        Check(Enabled(dlg, IDOK), "a buffer size that does not parse lights Apply, so Apply can say why");
        SetDlgItemTextW(dlg, IDC_ADV_BUF, buf);
        Check(!Enabled(dlg, IDOK), "putting the size back greys it");

        ShowPage(dlg, L"Output");
        Choose(dlg, IDC_OUT_FMT, 1);
        g_notes.clear();
        Press(dlg, IDOK);
    });
    Check(applied, "Apply reports the settings applied");
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "Apply applied JSON Lines");
    Check(Logged("options: applied -- FORMAT CSV -> JSONL"), "the log says what Apply changed, and nothing else",
          g_notes.empty() ? L"" : Wide(g_notes.back()));

    // ---- Cancel changes nothing -------------------------------------------------------------
    const bool cancelled = !Session("disarmed: Cancel", [](HWND dlg)
    {
        ShowPage(dlg, L"Output");
        Check(SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_GETCURSEL, 0, 0) == 1, "Format reopens on JSON Lines");
        Choose(dlg, IDC_OUT_FMT, 0);
        Press(dlg, IDCANCEL);
    });
    Check(cancelled, "Cancel reports nothing applied");
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "Cancel left the format as it was");

    // ---- armed: the format is locked, and only the log level can light Apply ----------------
    g_armed = true;
    g_notes.clear();
    Session("armed: the format is locked", [](HWND dlg)
    {
        Check(Logged("options: opened while armed -- XLL DEPTH="), "opening is logged with every setting",
              g_notes.empty() ? L"" : Wide(g_notes.front()));
        ShowPage(dlg, L"Output");
        Check(!Enabled(dlg, IDC_OUT_FMT), "Format is greyed while armed");
        Check(Visible(dlg, IDC_ARMEDNOTE) && Contains(TextOf(dlg, IDC_ARMEDNOTE), L"armed"),
              "Output says why", TextOf(dlg, IDC_ARMEDNOTE));
        Choose(dlg, IDC_OUT_FMT, 0);
        Check(!Enabled(dlg, IDOK), "a locked setting cannot light Apply");
        ShowPage(dlg, L"Advanced");
        Choose(dlg, IDC_ADV_LVL, 0);
        Check(Enabled(dlg, IDOK), "the log level can");
        Press(dlg, IDOK);
    });
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "Apply while armed did not change the format");
    Check(core::Log::GetLevel() == core::Log::Level::Debug, "Apply while armed changed the log level");
    Check(Logged("options: applied -- LOGLEVEL INFO -> DEBUG"), "and the log says only that");

    // ---- the breaks column: off by default, lights Apply, and is locked while armed ----------
    Session("armed: the breaks column is locked", [](HWND dlg)
    {
        ShowPage(dlg, L"Advanced");
        Check(!Enabled(dlg, IDC_ADV_BRK), "the breaks column is greyed while armed");
        Press(dlg, IDCANCEL);
    });
    g_armed = false;
    g_notes.clear();
    Session("disarmed: the breaks column", [](HWND dlg)
    {
        ShowPage(dlg, L"Advanced");
        Check(Enabled(dlg, IDC_ADV_BRK) && !Checked(dlg, IDC_ADV_BRK), "the breaks column is off, and can be turned on");
        Press(dlg, IDC_ADV_BRK);
        Check(Enabled(dlg, IDOK), "ticking it lights Apply");
        g_notes.clear();
        Press(dlg, IDOK);
    });
    Check(core::modes::GetBreakpoints(core::modes::Source::Vba), "Apply turned VBA BREAKPOINTS on");
    Check(Logged("options: applied -- VBA BREAKPOINTS FALSE -> TRUE"), "and the log says only that",
          g_notes.empty() ? L"" : Wide(g_notes.back()));
    core::modes::SetBreakpoints(core::modes::Source::Vba, false);

    std::printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
