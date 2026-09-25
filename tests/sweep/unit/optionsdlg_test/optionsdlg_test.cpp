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
#include "ui/eventpane.h"
#include "ui/ribbonmodel.h"
#include "app/settings.h"
#include "core/eventlist.h"
#include "core/tracemodes.h"
#include "core/log.h"

#include <windows.h>
#include <oleacc.h>
#include <atomic>
#include <climits>
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
    // This Excel lacks one event, so the page has a greyed row to show.
    namespace appevents
    {
        core::events::Mask Probe(bool& known)
        {
            known = true;
            return core::events::AllMask() & ~(core::events::Mask(1) << core::events::Find("RemoteWorkbookNewChart"));
        }
        core::events::Mask Available(bool& known) { return Probe(known); }
        bool ParamsText(int k, wchar_t* out, int cap)
        {
            if (k != core::events::Find("SheetChange")) return false;
            wcsncpy_s(out, cap, L"Sh As Object, Target As Range", _TRUNCATE);
            return true;
        }
    }
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

// The Capture page's settings, by control and by the model key that holds each.
struct CaptureSetting { int id; const wchar_t* key; bool depth; };
static const CaptureSetting kCapture[] = {
    { IDC_XLL_DEPTH, L"ddXllDepth", true },  { IDC_XLL_ARGS, L"cbXllArgs", false },
    { IDC_XLL_RET,   L"cbXllRet",   false }, { IDC_VBA_DEPTH, L"ddVbaDepth", true },
    { IDC_VBA_ARGS,  L"cbVbaArgs",  false }, { IDC_VBA_RET,  L"cbVbaRet",  false },
    { IDC_VBA_OBJ,   L"cbVbaObj",   false },
};
static int g_captureWas[sizeof(kCapture) / sizeof(kCapture[0])];
static int ModelValue(const CaptureSetting& c)
{
    return c.depth ? ui::ribbon::model::DepthIndex(c.key) : (ui::ribbon::model::ReadToggle(c.key) ? 1 : 0);
}

// What one action repaints, counted on the dialog's own thread as it paints.
static HWND g_countDlg = nullptr, g_countPane = nullptr;
static std::atomic<int> g_dlgErases{ 0 }, g_paneErases{ 0 }, g_boxDraws{ 0 };
static LRESULT CALLBACK CountPaints(int code, WPARAM wp, LPARAM lp)
{
    if (code == HC_ACTION)
    {
        const CWPSTRUCT* c = reinterpret_cast<const CWPSTRUCT*>(lp);
        if (c->message == WM_ERASEBKGND && c->hwnd == g_countDlg)  ++g_dlgErases;
        if (c->message == WM_ERASEBKGND && c->hwnd == g_countPane) ++g_paneErases;
        if (c->message == WM_DRAWITEM && c->hwnd == g_countPane)   ++g_boxDraws;
    }
    return CallNextHookEx(nullptr, code, wp, lp);
}
// WM_PAINT is not sent, so it is seen as the dialog's loop takes it from the queue.
static HWND g_countCombo = nullptr;
static std::atomic<int> g_comboPaints{ 0 };
static LRESULT CALLBACK CountQueuedPaints(int code, WPARAM wp, LPARAM lp)
{
    if (code == HC_ACTION && wp == PM_REMOVE)
    {
        const MSG* m = reinterpret_cast<const MSG*>(lp);
        if (m->message == WM_PAINT && m->hwnd == g_countCombo) ++g_comboPaints;
    }
    return CallNextHookEx(nullptr, code, wp, lp);
}
static void ResetCounts() { g_dlgErases = 0; g_paneErases = 0; g_boxDraws = 0; g_comboPaints = 0; }
// The dialog paints from its own message loop, once the action's messages are handled.
static void Settle() { Sleep(400); }

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

// The Events page's boxes live in its pane, so a click is sent there.
static HWND Pane(HWND dlg) { return GetDlgItem(dlg, IDC_EVT_PANE); }
static int EventId(const char* name) { return ui::eventpane::kFirstEventId + core::events::Find(name); }
static int GroupId(core::events::Group g) { return ui::eventpane::kFirstGroupId + static_cast<int>(g); }
static void Tick(HWND dlg, int id)
{
    HWND pane = Pane(dlg);
    SendMessageW(pane, WM_COMMAND, MAKEWPARAM(id, BN_CLICKED), reinterpret_cast<LPARAM>(GetDlgItem(pane, id)));
}
static bool On(HWND dlg, const char* name)
{
    return ((ui::eventpane::Selected(Pane(dlg)) >> core::events::Find(name)) & 1) != 0;
}
static int Preset(HWND dlg) { return static_cast<int>(SendDlgItemMessageW(dlg, IDC_EVT_PRESET, CB_GETCURSEL, 0, 0)); }

// What a screen reader is told about a box.
static long AccState(HWND box)
{
    IAccessible* acc = nullptr;
    if (FAILED(AccessibleObjectFromWindow(box, static_cast<DWORD>(OBJID_CLIENT), IID_IAccessible,
                                          reinterpret_cast<void**>(&acc))) || !acc) return -1;
    VARIANT self; VariantInit(&self); self.vt = VT_I4; self.lVal = CHILDID_SELF;
    VARIANT st; VariantInit(&st);
    long out = -1;
    if (SUCCEEDED(acc->get_accState(self, &st)) && st.vt == VT_I4) out = st.lVal;
    acc->Release();
    return out;
}

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
    // Excel's main thread has COM; the dialog's screen-reader annotations need it here too.
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
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
        Check(names == L"Capture,Events,Output,Advanced,About,Notices", "the categories are Capture, Events, Output, Advanced, About, Notices", names);

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

    // ---- the Capture page: each setting shows the model's value, and Apply writes it back -------
    for (size_t i = 0; i < sizeof(kCapture) / sizeof(kCapture[0]); ++i) g_captureWas[i] = ModelValue(kCapture[i]);
    g_notes.clear();
    const bool captureApplied = Session("disarmed: the Capture page", [](HWND dlg)
    {
        ShowPage(dlg, L"Capture");
        for (size_t i = 0; i < sizeof(kCapture) / sizeof(kCapture[0]); ++i)
        {
            const CaptureSetting& c = kCapture[i];
            const int shown = c.depth ? static_cast<int>(SendDlgItemMessageW(dlg, c.id, CB_GETCURSEL, 0, 0))
                                      : (Checked(dlg, c.id) ? 1 : 0);
            Check(shown == g_captureWas[i] && Enabled(dlg, c.id), "a Capture setting shows the model's value", c.key);
        }
        for (size_t i = 0; i < sizeof(kCapture) / sizeof(kCapture[0]); ++i)
        {
            const CaptureSetting& c = kCapture[i];
            if (c.depth) Choose(dlg, c.id, g_captureWas[i] == 0 ? 1 : 0);
            else         Press(dlg, c.id);
        }
        Check(Enabled(dlg, IDOK), "changing them lights Apply");
        Press(dlg, IDOK);
    });
    Check(captureApplied, "Apply reports the Capture page applied");
    for (size_t i = 0; i < sizeof(kCapture) / sizeof(kCapture[0]); ++i)
    {
        const CaptureSetting& c = kCapture[i];
        const int want = c.depth ? (g_captureWas[i] == 0 ? 1 : 0) : 1 - g_captureWas[i];
        Check(ModelValue(c) == want, "Apply wrote a Capture setting", c.key);
        if (c.depth) ui::ribbon::model::SetDepthIndex(c.key, g_captureWas[i]);
        else         ui::ribbon::model::WriteToggle(c.key, g_captureWas[i] != 0);
    }
    g_armed = true;
    Session("armed: the Capture page is locked", [](HWND dlg)
    {
        ShowPage(dlg, L"Capture");
        for (const CaptureSetting& c : kCapture)
            Check(!Enabled(dlg, c.id), "a Capture setting is greyed while armed", c.key);
        Press(dlg, IDCANCEL);
    });
    g_armed = false;

    // ---- the Events page: presets, ticks, the heading cycle, and what a screen reader hears ------
    g_notes.clear();
    const bool eventsApplied = Session("disarmed: the Events page", [](HWND dlg)
    {
        CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
        using core::events::Group;
        ShowPage(dlg, L"Events");
        Check(Visible(dlg, IDC_EVT_PRESET) && Visible(dlg, IDC_EVT_PANE), "Events shows the preset and the list");
        Check(ComboItem(dlg, IDC_EVT_PRESET, 0) == L"None" && ComboItem(dlg, IDC_EVT_PRESET, 1) == L"Calc" &&
              ComboItem(dlg, IDC_EVT_PRESET, 2) == L"Selection" && ComboItem(dlg, IDC_EVT_PRESET, 3) == L"Calc & Selection" &&
              ComboItem(dlg, IDC_EVT_PRESET, 4) == L"All" && ComboItem(dlg, IDC_EVT_PRESET, 5) == L"Custom",
              "the presets are None, Calc, Selection, Calc & Selection, All, Custom");
        Check(Preset(dlg) == 1, "the default reads Calc");
        Check(TextOf(dlg, IDC_EVT_COUNT) == L"9 of 53 events", "the count leaves out the event this Excel lacks",
              TextOf(dlg, IDC_EVT_COUNT));
        HWND absent = GetDlgItem(Pane(dlg), EventId("RemoteWorkbookNewChart"));
        Check(absent && !IsWindowEnabled(absent), "an event this Excel lacks is greyed");
        Check(ui::eventpane::TipOf(Pane(dlg), EventId("SheetChange")) == L"Records: Sh As Object, Target As Range",
              "hovering an event shows the parameters it records", ui::eventpane::TipOf(Pane(dlg), EventId("SheetChange")));

        Tick(dlg, EventId("SheetSelectionChange"));
        Check(On(dlg, "SheetSelectionChange") && Preset(dlg) == 5, "a single tick turns the preset to Custom");
        Check(TextOf(dlg, IDC_EVT_COUNT) == L"10 of 53 events", "and the count follows", TextOf(dlg, IDC_EVT_COUNT));
        Check(Enabled(dlg, IDOK), "a changed tick lights Apply");
        Tick(dlg, EventId("SheetSelectionChange"));
        Check(Preset(dlg) == 1 && !Enabled(dlg, IDOK), "unticking it again reads Calc, and greys Apply");

        Choose(dlg, IDC_EVT_PRESET, 2);
        Check(On(dlg, "SheetSelectionChange") && !On(dlg, "SheetChange") && Preset(dlg) == 2,
              "choosing Selection ticks exactly its events");
        Choose(dlg, IDC_EVT_PRESET, 3);
        Check(On(dlg, "SheetSelectionChange") && On(dlg, "SheetChange") && !On(dlg, "WindowResize") &&
              Preset(dlg) == 3, "choosing Calc & Selection ticks both sets and nothing else");
        Choose(dlg, IDC_EVT_PRESET, 1);
        for (int k = 0; k < core::events::Count(); ++k)
            if (core::events::At(k).selection && !On(dlg, core::events::At(k).name))
                Tick(dlg, ui::eventpane::kFirstEventId + k);
        Check(Preset(dlg) == 3, "ticking Selection's events onto Calc reads Calc & Selection");
        Choose(dlg, IDC_EVT_PRESET, 1);

        // PivotTables under Calc is mixed: two of its seven events.
        const int pivots = GroupId(Group::PivotTables);
        HWND pane = Pane(dlg);
        Check(ui::eventpane::GroupState(pane, static_cast<int>(Group::PivotTables)) == 2, "a partly ticked group is mixed");
        Check(ui::eventpane::TipOf(pane, pivots) == L"Click to tick all 7", "a mixed heading offers to tick all",
              ui::eventpane::TipOf(pane, pivots));
        Tick(dlg, pivots);
        Check(ui::eventpane::GroupState(pane, static_cast<int>(Group::PivotTables)) == 1 && Preset(dlg) == 5,
              "the first click ticks the whole group");
        Check(ui::eventpane::TipOf(pane, pivots) == L"Click to clear all", "then offers to clear it");
        Tick(dlg, pivots);
        Check(ui::eventpane::GroupState(pane, static_cast<int>(Group::PivotTables)) == 0, "the second click clears it");
        Check(ui::eventpane::TipOf(pane, pivots) == L"Click to restore your selection", "then offers the mix back");
        Tick(dlg, pivots);
        Check(Preset(dlg) == 1 && On(dlg, "SheetPivotTableUpdate") && !On(dlg, "WorkbookPivotTableOpenConnection"),
              "the third click restores exactly the mix it had");

        // A single tick in the group forgets the mix.
        Tick(dlg, pivots); Tick(dlg, pivots);
        Tick(dlg, EventId("WorkbookPivotTableOpenConnection"));
        Tick(dlg, EventId("WorkbookPivotTableOpenConnection"));
        Check(ui::eventpane::TipOf(pane, pivots) == L"Click to tick all 7", "a single tick forgets the remembered mix",
              ui::eventpane::TipOf(pane, pivots));
        // So does a preset.
        Choose(dlg, IDC_EVT_PRESET, 1);
        Tick(dlg, pivots); Tick(dlg, pivots);
        Choose(dlg, IDC_EVT_PRESET, 0);
        Check(ui::eventpane::TipOf(pane, pivots) == L"Click to tick all 7", "choosing a preset forgets it too",
              ui::eventpane::TipOf(pane, pivots));

        Choose(dlg, IDC_EVT_PRESET, 1);
        Tick(dlg, EventId("WindowResize"));
        HWND calc = GetDlgItem(pane, GroupId(Group::CalcData));
        const long mixed = AccState(GetDlgItem(pane, GroupId(Group::PivotTables)));
        Check(AccState(GetDlgItem(pane, EventId("WindowResize"))) != -1 &&
              (AccState(GetDlgItem(pane, EventId("WindowResize"))) & STATE_SYSTEM_CHECKED) != 0,
              "a screen reader hears a ticked event as checked");
        Check((AccState(calc) & STATE_SYSTEM_CHECKED) != 0 && (mixed & STATE_SYSTEM_MIXED) != 0,
              "and a heading as checked or mixed");
        g_notes.clear();
        Press(dlg, IDOK);
        CoUninitialize();
    });
    Check(eventsApplied, "Apply reports the events applied");
    Check(((core::events::GetSelected() >> core::events::Find("WindowResize")) & 1) != 0, "Apply ticked WindowResize");
    Check(Logged("options: EVENTS WindowResize -> TRUE"), "and logged that one event");

    g_armed = true;
    Session("armed: the Events page is locked", [](HWND dlg)
    {
        ShowPage(dlg, L"Events");
        Check(!Enabled(dlg, IDC_EVT_PRESET), "the preset is greyed while armed");
        HWND box = GetDlgItem(Pane(dlg), EventId("SheetChange"));
        Check(box && !IsWindowEnabled(box), "and so are the boxes");
        Tick(dlg, EventId("SheetChange"));
        Check(!Enabled(dlg, IDOK), "a click while armed changes nothing");
        Press(dlg, IDCANCEL);
    });
    g_armed = false;

    // ---- painting: an action redraws what it changed, and the dialog does not flash ----------
    Session("disarmed: what an action repaints", [](HWND dlg)
    {
        g_countDlg = dlg; g_countPane = Pane(dlg);
        HHOOK hook = SetWindowsHookExW(WH_CALLWNDPROC, CountPaints, nullptr, GetWindowThreadProcessId(dlg, nullptr));
        Check(hook != nullptr, "the paint counter is installed");
        ShowPage(dlg, L"Events"); Settle();

        ResetCounts(); ShowPage(dlg, L"Output"); Settle();
        const int pageErases = g_dlgErases;
        ResetCounts(); ShowPage(dlg, L"Events"); Settle();
        const int backErases = g_dlgErases;
        std::printf("   page switch: dialog erased %d then %d times\n", pageErases, backErases);
        Check(pageErases <= 1 && backErases <= 1, "a page switch erases the dialog once at most");

        // The first group's rows are in view, so their redraws are seen.
        ResetCounts(); Tick(dlg, EventId("SheetChange")); Settle();
        const int tickDraws = g_boxDraws;
        Tick(dlg, EventId("SheetChange")); Settle();
        std::printf("   one tick: %d boxes drawn\n", tickDraws);
        Check(tickDraws >= 1 && tickDraws <= 2, "a tick redraws its box and its heading, nothing else");

        ResetCounts(); SendMessageW(Pane(dlg), WM_VSCROLL, SB_LINEDOWN, 0); Settle();
        const int scrollDraws = g_boxDraws, scrollErases = g_paneErases;
        std::printf("   one line of scroll: %d boxes drawn, pane erased %d times\n", scrollDraws, scrollErases);
        Check(scrollDraws <= 3, "a line of scroll draws only the rows it uncovers");

        // Every row, in view or not, sits where the rows before it put it, however far it scrolled.
        {
            HWND pane = Pane(dlg);
            for (int i = 0; i < 12; ++i) SendMessageW(pane, WM_VSCROLL, SB_LINEDOWN, 0);
            SendMessageW(pane, WM_VSCROLL, SB_LINEUP, 0);
            for (int i = 0; i < 5; ++i) SendMessageW(pane, WM_VSCROLL, SB_PAGEDOWN, 0);
            SendMessageW(pane, WM_VSCROLL, SB_TOP, 0);
            SendMessageW(pane, WM_VSCROLL, SB_PAGEDOWN, 0);
            int misplaced = 0, expected = INT_MIN;
            for (HWND c = GetWindow(pane, GW_CHILD); c; c = GetWindow(c, GW_HWNDNEXT))
            {
                RECT r; GetWindowRect(c, &r);
                MapWindowPoints(nullptr, pane, reinterpret_cast<POINT*>(&r), 2);
                if (expected != INT_MIN && r.top != expected) ++misplaced;
                expected = r.bottom;
            }
            std::printf("   after scrolling: %d rows out of place\n", misplaced);
            Check(misplaced == 0, "rows scrolled out of view come back in place");
            SendMessageW(pane, WM_VSCROLL, SB_TOP, 0);
        }

        ResetCounts(); Choose(dlg, IDC_EVT_PRESET, 4); Settle();
        const int presetDraws = g_boxDraws;
        std::printf("   choosing All: %d boxes drawn\n", presetDraws);

        // A tick that changes the preset shows it at once, not at the next hover.
        g_countCombo = GetDlgItem(dlg, IDC_EVT_PRESET);
        HHOOK queued = SetWindowsHookExW(WH_GETMESSAGE, CountQueuedPaints, nullptr, GetWindowThreadProcessId(dlg, nullptr));
        Choose(dlg, IDC_EVT_PRESET, 3); Settle();
        ResetCounts(); Tick(dlg, EventId("SheetSelectionChange")); Settle();
        const int toCustom = g_comboPaints;
        ResetCounts(); Tick(dlg, EventId("SheetSelectionChange")); Settle();
        const int toBoth = g_comboPaints;
        std::printf("   the preset repainted %d then %d times\n", toCustom, toBoth);
        Check(Preset(dlg) == 3 && toCustom >= 1 && toBoth >= 1,
              "the drop-down repaints when a tick turns it to Custom and back");
        if (queued) UnhookWindowsHookEx(queued);
        Choose(dlg, IDC_EVT_PRESET, 1); Settle();

        if (hook) UnhookWindowsHookEx(hook);
        g_countDlg = g_countPane = g_countCombo = nullptr;
        Press(dlg, IDCANCEL);
    });

    // ---- the mouse and keyboard as a user drives them ---------------------------------------
    Session("disarmed: quick clicks and browsing the presets", [](HWND dlg)
    {
        // Two clicks in quick succession: an owner-drawn button reports the second as a double-click.
        auto twice = [](HWND box) {
            const LPARAM at = MAKELPARAM(4, 4);
            PostMessageW(box, WM_LBUTTONDOWN, MK_LBUTTON, at); PostMessageW(box, WM_LBUTTONUP, 0, at);
            PostMessageW(box, WM_LBUTTONDBLCLK, MK_LBUTTON, at); PostMessageW(box, WM_LBUTTONUP, 0, at);
            Settle(); };
        const bool args = Checked(dlg, IDC_XLL_ARGS);
        twice(GetDlgItem(dlg, IDC_XLL_ARGS));
        Check(Checked(dlg, IDC_XLL_ARGS) == args, "two quick clicks on a box tick and untick it");
        ShowPage(dlg, L"Events"); Settle();
        const bool change = On(dlg, "SheetChange");
        twice(GetDlgItem(Pane(dlg), EventId("SheetChange")));
        Check(On(dlg, "SheetChange") == change, "and on an event's box");

        // Arrowing through the open list only looks; Esc leaves the ticks as they were.
        Choose(dlg, IDC_EVT_PRESET, 1);
        Tick(dlg, EventId("WindowResize"));
        const std::wstring before = TextOf(dlg, IDC_EVT_COUNT);
        HWND combo = GetDlgItem(dlg, IDC_EVT_PRESET);
        SetFocus(combo);
        PostMessageW(combo, CB_SHOWDROPDOWN, TRUE, 0); Settle();
        PostMessageW(combo, WM_KEYDOWN, VK_UP, 0); Settle();
        Check(TextOf(dlg, IDC_EVT_COUNT) == before && On(dlg, "WindowResize"),
              "arrowing through the open list changes no tick", TextOf(dlg, IDC_EVT_COUNT));
        PostMessageW(combo, WM_KEYDOWN, VK_ESCAPE, 0); Settle();
        Check(TextOf(dlg, IDC_EVT_COUNT) == before && Preset(dlg) == 5,
              "and Esc puts the list back on Custom", TextOf(dlg, IDC_EVT_COUNT));
        // Enter on a preset applies it.
        PostMessageW(combo, CB_SHOWDROPDOWN, TRUE, 0); Settle();
        PostMessageW(combo, WM_KEYDOWN, VK_UP, 0); Settle();
        PostMessageW(combo, WM_KEYDOWN, VK_RETURN, 0); Settle();
        Check(Preset(dlg) == 4 && TextOf(dlg, IDC_EVT_COUNT) == L"53 of 53 events",
              "Enter on All applies it", TextOf(dlg, IDC_EVT_COUNT));
        Press(dlg, IDCANCEL);
    });

    std::printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
