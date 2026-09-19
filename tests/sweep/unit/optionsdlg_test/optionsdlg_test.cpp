// The Options dialog, opened for real with no Excel: the shipping dialog code and resources, a
// second thread driving it through its controls, and stubs for what the add-in's session would
// supply. Plus the text reflow the About and Notices pages use, on its own.
//
// Excel is not needed and not wanted: the harness's dialog watchdog would click away a dialog
// shown inside a test's Excel before it could be driven.
#include "ui/optionsdlg.h"
#include "ui/optionsres.h"
#include "ui/reflow.h"
#include "core/tracemodes.h"
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
    std::wstring EnsureAppSubdir(const wchar_t* leaf) { return std::wstring(L"C:\\stub\\") + leaf; }
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

    core::modes::SetFormat(core::modes::Format::Csv);

    // ---- disarmed: every page, the format moved to Output, OK applies it --------------------
    const bool applied = Session("disarmed: pages, texts, format, OK", [](HWND dlg)
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
        Check(TextOf(dlg, IDC_ADV_SEC1) == L"Buffer", "Advanced's first section is Buffer", TextOf(dlg, IDC_ADV_SEC1));

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

        ShowPage(dlg, L"Output");
        SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_SETCURSEL, 1, 0);
        Press(dlg, IDOK);
    });
    Check(applied, "OK reports the settings applied");
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "OK applied JSON Lines");

    // ---- Cancel changes nothing -------------------------------------------------------------
    const bool cancelled = !Session("disarmed: Cancel", [](HWND dlg)
    {
        ShowPage(dlg, L"Output");
        Check(SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_GETCURSEL, 0, 0) == 1, "Format reopens on JSON Lines");
        SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_SETCURSEL, 0, 0);
        Press(dlg, IDCANCEL);
    });
    Check(cancelled, "Cancel reports nothing applied");
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "Cancel left the format as it was");

    // ---- armed: the format is locked, and OK cannot change it -------------------------------
    g_armed = true;
    Session("armed: the format is locked", [](HWND dlg)
    {
        ShowPage(dlg, L"Output");
        Check(!Enabled(dlg, IDC_OUT_FMT), "Format is greyed while armed");
        Check(Visible(dlg, IDC_ARMEDNOTE) && Contains(TextOf(dlg, IDC_ARMEDNOTE), L"armed"),
              "Output says why", TextOf(dlg, IDC_ARMEDNOTE));
        SendDlgItemMessageW(dlg, IDC_OUT_FMT, CB_SETCURSEL, 0, 0);
        Press(dlg, IDOK);
    });
    Check(core::modes::GetFormat() == core::modes::Format::Jsonl, "OK while armed did not change the format");
    g_armed = false;

    std::printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
