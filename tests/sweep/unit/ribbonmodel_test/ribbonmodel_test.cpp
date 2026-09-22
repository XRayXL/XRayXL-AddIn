// The ribbon's decisions, without Excel or COM: above all, the XML against the handlers, both ways.
#include "ui/ribbonmodel.h"
#include "core/tracemodes.h"

#include <cstdio>
#include <cwchar>
#include <string>
#include <vector>

namespace M = ui::ribbon::model;
using core::modes::Depth;
using core::modes::Source;

static int g_fail = 0;
static int g_pass = 0;

static void Check(const char* name, bool ok, const char* detail = "")
{
    if (ok) { ++g_pass; std::printf("  ok   %s\n", name); }
    else    { ++g_fail; std::printf("  FAIL %s  %s\n", name, detail); }
}

// Every value of attribute `attr`, by a plain scan: this checks the literal string that ships.
static std::vector<std::wstring> AttrValues(const std::wstring& xml, const wchar_t* attr)
{
    std::vector<std::wstring> out;
    const std::wstring needle = std::wstring(attr) + L"='";
    size_t at = 0;
    while ((at = xml.find(needle, at)) != std::wstring::npos)
    {
        const size_t from = at + needle.size();
        const size_t to = xml.find(L'\'', from);
        if (to == std::wstring::npos) break;
        out.push_back(xml.substr(from, to - from));
        at = to;
    }
    return out;
}

static std::string Narrow(const std::wstring& w)
{
    std::string s;
    for (wchar_t c : w) s += (c < 128) ? static_cast<char>(c) : '?';
    return s;
}

int main()
{
    const std::wstring xml = M::kCustomUi;
    std::printf("ribbonmodel_test\n");

    // ---- the XML itself ---------------------------------------------------
    Check("xml-uses-customui14",
          xml.find(L"http://schemas.microsoft.com/office/2009/07/customui") != std::wstring::npos);
    Check("xml-is-balanced",
          xml.rfind(L"<customUI", 0) == 0 && xml.find(L"</customUI>") == xml.size() - 11);
    // size='large' needs an image, and an imageMso Office does not know fails silently: render it first.
    const wchar_t* kRenderedOnRealExcel[] = {
        L"MacroRecord", L"MacroPlay", L"TraceDependents", L"FileSave", L"FileClose",
        L"Delete", L"Refresh", L"ClearAll", L"CancelRequest", L"DeclineInvitation",
        L"PauseTimer", L"HighImportance", L"AnimationCustom", L"SaveAll",
        L"FileSaveAsExcelXlsx", L"TraceError", L"ApplicationOptionsDialog",
        L"FileDocumentInspect",
    };
    for (const std::wstring& img : AttrValues(xml, L"imageMso"))
    {
        bool known = false;
        for (const wchar_t* k : kRenderedOnRealExcel) if (img == k) { known = true; break; }
        Check(("imageMso-is-verified:" + Narrow(img)).c_str(), known,
              "not on the rendered-on-real-Excel list -- render it before shipping it: "
              "an id Office does not know degrades the button to small text, silently");
    }

    for (const std::wstring& id : { std::wstring(L"btnArm"), std::wstring(L"btnDisarm"),
                                    std::wstring(L"btnOptions"), std::wstring(L"btnDiagnostics") })
    {
        const size_t at = xml.find(L"id='" + id + L"'");
        const size_t end = xml.find(L"/>", at);
        const std::wstring tag = (at == std::wstring::npos) ? L"" : xml.substr(at, end - at);
        const bool large = tag.find(L"size='large'") != std::wstring::npos;
        const bool image = tag.find(L"imageMso='") != std::wstring::npos ||
                           tag.find(L"image='") != std::wstring::npos;
        Check(("button-is-full-size:" + Narrow(id)).c_str(), large && image,
              "needs size='large' AND an image -- size alone renders a blank tall "
              "button with a shrunken label");
    }
    // a wrong built-in tab id makes Office reject the whole customUI
    Check("hosted-on-the-developer-tab",
          xml.find(L"idMso='TabDeveloper'") != std::wstring::npos,
          "the button lives on Excel's Developer tab, not a tab of our own");
    Check("no-tab-of-our-own", xml.find(L"<tab id='") == std::wstring::npos,
          "a tab with an id of ours would be a second home for the same thing");
    // Four items, in the order a person reads them.
    const size_t pArm = xml.find(L"id='btnArm'");
    const size_t pDis = xml.find(L"id='btnDisarm'");
    const size_t pOpt = xml.find(L"id='btnOptions'");
    const size_t pDiag = xml.find(L"id='btnDiagnostics'");
    Check("menu-has-arm-disarm-options-diagnostics",
          pArm != std::wstring::npos && pDis != std::wstring::npos &&
          pOpt != std::wstring::npos && pDiag != std::wstring::npos &&
          pArm < pDis && pDis < pOpt && pOpt < pDiag,
          "Arm, Disarm, Options, Diagnostics -- in that order");

    Check("xml-declares-onload", xml.find(L"onLoad='OnRibbonLoad'") != std::wstring::npos,
          "without onLoad there is no IRibbonUI, so nothing can ever be invalidated");

    // ---- every callback the XML names must resolve ------------------------
    const wchar_t* kCallbackAttrs[] = {
        L"onLoad", L"loadImage", L"onAction", L"getEnabled", L"getPressed",
        L"getSelectedItemIndex", L"getLabel"
    };
    int callbacksSeen = 0;
    for (const wchar_t* attr : kCallbackAttrs)
    {
        for (const std::wstring& name : AttrValues(xml, attr))
        {
            ++callbacksSeen;
            const bool known = M::CallbackForName(name.c_str()) != M::CbUnknown;
            Check(("xml-callback-resolves:" + Narrow(name)).c_str(), known,
                  "named in the XML but GetIDsOfNames would refuse it -- a dead control");
        }
    }
    // a floor, so a scan that stops finding anything cannot pass
    Check("xml-actually-named-callbacks", callbacksSeen >= 7, "the scan found suspiciously few");

    // a name that is not a callback must be refused, or the check above proves nothing
    Check("unknown-callback-refused", M::CallbackForName(L"OnSomethingElse") == M::CbUnknown);
    Check("null-callback-refused",    M::CallbackForName(nullptr) == M::CbUnknown);
    // Office asks by the name in the XML; it has been seen to vary in case.
    Check("callback-lookup-is-case-insensitive",
          M::CallbackForName(L"getenabled") == M::CbGetEnabled);

    // ---- every control id in the XML must be handled ----
    const std::vector<std::wstring> kNotControls = { L"grpXRay" };
    std::vector<std::wstring> controls;
    for (const std::wstring& id : AttrValues(xml, L"id"))
    {
        bool skip = false;
        for (const std::wstring& n : kNotControls) if (n == id) { skip = true; break; }
        if (skip) continue;
        controls.push_back(id);
        Check(("xml-control-is-handled:" + Narrow(id)).c_str(),
              M::KnownControl(id.c_str()),
              "in the XML but no handler knows it -- the control would do nothing");
    }
    // a bare count: a new control must be given a handler
    Check("xml-has-every-control", controls.size() == 4,
          "expected 4: Arm, Disarm, Options and Diagnostics");

    // ...and every handled control must be in the XML
    const wchar_t* kExpected[] = { L"btnArm", L"btnDisarm", L"btnOptions", L"btnDiagnostics" };
    for (const wchar_t* id : kExpected)
    {
        bool inXml = false;
        for (const std::wstring& c : controls) if (c == id) { inXml = true; break; }
        Check(("handled-control-is-in-the-xml:" + Narrow(id)).c_str(), inXml,
              "handled here but absent from the XML -- unreachable");
    }
    Check("unknown-control-is-not-claimed", !M::KnownControl(L"btnNoSuchThing"));

    // ---- enablement: the command surface's rule, shown ---------------------
    Check("not-armed-arm-enabled",      M::EnabledFor(L"btnArm", false));
    Check("not-armed-disarm-disabled", !M::EnabledFor(L"btnDisarm", false));
    Check("armed-arm-disabled",        !M::EnabledFor(L"btnArm", true));
    Check("armed-disarm-enabled",       M::EnabledFor(L"btnDisarm", true));

    // Options is always live: the dialog greys what cannot be changed.
    // Diagnostics is always live too: it reads, and changes nothing.
    for (int a = 0; a < 2; ++a)
    {
        Check("options-is-always-live", M::EnabledFor(L"btnOptions", a != 0));
        Check("diagnostics-is-always-live", M::EnabledFor(L"btnDiagnostics", a != 0));
    }
    // the settings carry the setters' rule: refused while armed
    for (const std::wstring& id : { std::wstring(L"cbXllArgs"), std::wstring(L"cbVbaObj"),
                                    std::wstring(L"ddXllDepth"), std::wstring(L"cbPauseFull") })
    {
        Check(("setting-greyed-while-armed:" + Narrow(id)).c_str(),
              !M::EnabledFor(id.c_str(), true) && M::EnabledFor(id.c_str(), false),
              "the setters refuse while armed, so the control must be greyed then");
    }

    // OBJECTS is VBA only: an XLL argument has no object model to describe.
    Check("xll-has-no-objects-control", !M::KnownControl(L"cbXllObj"),
          "OBJECTS is not honoured by src/xll -- offering it promises nothing");
    Check("xll-has-no-breakpoints-control", !M::KnownControl(L"cbXllBrk"),
          "an XLL has no breakpoints");

    // ---- the buffer size box: SetTraceParam's own words and parser ----
    {
        wchar_t buf[32];
        core::modes::SetBufferBytes(64ull * 1024 * 1024);
        M::BufferText(buf, 32);
        Check("buffer-text-renders-mb", wcscmp(buf, L"64MB") == 0, Narrow(buf).c_str());

        Check("buffer-takes-mb", M::SetBufferText(L"128MB"));
        Check("buffer-mb-took-effect", core::modes::GetBufferBytes() == 128ull * 1024 * 1024);

        Check("buffer-takes-kb", M::SetBufferText(L"512KB"));
        Check("buffer-kb-took-effect", core::modes::GetBufferBytes() == 512ull * 1024);
        M::BufferText(buf, 32);
        Check("buffer-text-renders-kb", wcscmp(buf, L"512KB") == 0, Narrow(buf).c_str());

        // Bare number is megabytes, and lower case is the same word.
        Check("buffer-bare-number-is-mb", M::SetBufferText(L"32"));
        Check("buffer-bare-took-effect", core::modes::GetBufferBytes() == 32ull * 1024 * 1024);
        Check("buffer-lower-case", M::SetBufferText(L"8mb"));
        Check("buffer-lower-case-took-effect", core::modes::GetBufferBytes() == 8ull * 1024 * 1024);

        // 0 is synchronous, and is the ONLY value below the ring floor.
        Check("buffer-zero-is-synchronous", M::SetBufferText(L"0"));
        Check("buffer-zero-took-effect", core::modes::GetBufferBytes() == 0);
        M::BufferText(buf, 32);
        Check("buffer-zero-renders-as-0", wcscmp(buf, L"0") == 0, Narrow(buf).c_str());

        // refused, and nothing changed
        core::modes::SetBufferBytes(64ull * 1024 * 1024);
        const std::size_t before = core::modes::GetBufferBytes();
        Check("buffer-refuses-below-the-floor", !M::SetBufferText(L"1KB"));
        Check("buffer-refuses-nonsense",        !M::SetBufferText(L"banana"));
        Check("buffer-refuses-bad-unit",        !M::SetBufferText(L"10GB"));
        Check("buffer-refuses-over-the-cap",    !M::SetBufferText(L"8192MB"));
        Check("buffer-refuses-empty",           !M::SetBufferText(L""));
        Check("buffer-refusals-changed-nothing", core::modes::GetBufferBytes() == before);

        // Symmetric: what the box shows, the box accepts back.
        core::modes::SetBufferBytes(128ull * 1024 * 1024);
        M::BufferText(buf, 32);
        Check("buffer-round-trips", M::SetBufferText(buf) &&
              core::modes::GetBufferBytes() == 128ull * 1024 * 1024, Narrow(buf).c_str());
    }

    // ---- check boxes read and write the real settings ----------------------
    const wchar_t* kToggleIds[] = {
        L"cbXllArgs", L"cbXllRet",
        L"cbVbaArgs", L"cbVbaRet", L"cbVbaObj", L"cbVbaBrk", L"cbPauseFull"
    };
    for (const wchar_t* id : kToggleIds)
    {
        Check(("is-a-toggle:" + Narrow(id)).c_str(), M::IsToggle(id));
        M::WriteToggle(id, false);
        const bool offOk = !M::ReadToggle(id);
        M::WriteToggle(id, true);
        const bool onOk = M::ReadToggle(id);
        Check(("toggle-round-trips:" + Narrow(id)).c_str(), offOk && onOk,
              "the check box does not read back what it set");
    }
    // All off, then one on, must leave exactly that one on: catches two ids wired to one setting.
    for (const wchar_t* id : kToggleIds)
    {
        for (const wchar_t* other : kToggleIds) M::WriteToggle(other, false);
        M::WriteToggle(id, true);
        int on = 0;
        for (const wchar_t* other : kToggleIds) if (M::ReadToggle(other)) ++on;
        Check(("toggle-is-its-own-setting:" + Narrow(id)).c_str(), on == 1,
              "more than one control changed, so two ids share a setting");
    }
    Check("non-toggle-reads-false", !M::ReadToggle(L"btnArm"));
    Check("non-toggle-is-not-a-toggle", !M::IsToggle(L"ddXllDepth"));

    // ---- the depth drop-downs: the item index is the Depth enum value ----
    Check("enum-order-is-off-top-all",
          static_cast<int>(Depth::Off) == 0 && static_cast<int>(Depth::Top) == 1 &&
          static_cast<int>(Depth::All) == 2);

    struct { const wchar_t* id; Source src; } kDepths[] = {
        { L"ddXllDepth", Source::Xll }, { L"ddVbaDepth", Source::Vba }
    };
    for (const auto& d : kDepths)
    {
        Source got = Source::Vba;
        Check(("is-a-depth-control:" + Narrow(d.id)).c_str(),
              M::IsDepthControl(d.id, got) && got == d.src,
              "the drop-down is wired to the wrong source");
        for (int ix = 0; ix <= 2; ++ix)
        {
            Check(("depth-set-accepted:" + Narrow(d.id) + ":" + std::to_string(ix)).c_str(),
                  M::SetDepthIndex(d.id, ix));
            Check(("depth-round-trips:" + Narrow(d.id) + ":" + std::to_string(ix)).c_str(),
                  M::DepthIndex(d.id) == ix &&
                  static_cast<int>(core::modes::GetDepth(d.src)) == ix,
                  "the drop-down and core::modes disagree");
        }
        // out of range changes nothing; it does not clamp
        M::SetDepthIndex(d.id, 2);
        Check(("depth-refuses-out-of-range:" + Narrow(d.id)).c_str(),
              !M::SetDepthIndex(d.id, 3) && !M::SetDepthIndex(d.id, -1) &&
              M::DepthIndex(d.id) == 2,
              "an index outside the list changed the setting");
    }
    // The two sources are independent -- one drop-down must not move the other.
    M::SetDepthIndex(L"ddXllDepth", 0);
    M::SetDepthIndex(L"ddVbaDepth", 2);
    Check("depth-sources-are-independent",
          M::DepthIndex(L"ddXllDepth") == 0 && M::DepthIndex(L"ddVbaDepth") == 2);
    Check("non-depth-control-has-no-index", M::DepthIndex(L"cbXllArgs") == -1);
    Check("non-depth-set-refused", !M::SetDepthIndex(L"cbXllArgs", 1));

    std::printf("\n%d passed, %d failed\n", g_pass, g_fail);
    return g_fail == 0 ? 0 : 1;
}
