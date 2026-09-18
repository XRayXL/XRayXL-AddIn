#include "ribbonmodel.h"
#include "../app/paramparse.h"
#include <cwctype>

#include <cwchar>

using core::modes::Depth;
using core::modes::Param;
using core::modes::Source;

namespace ui
{
namespace ribbon
{
namespace model
{
namespace
{
    struct Toggle { const wchar_t* id; Source src; Param param; };

    // The per-source check boxes. OBJECTS is VBA only.
    const Toggle kToggles[] = {
        { L"cbXllArgs", Source::Xll, Param::Args    },
        { L"cbXllRet",  Source::Xll, Param::RetVal  },
        { L"cbVbaArgs", Source::Vba, Param::Args    },
        { L"cbVbaRet",  Source::Vba, Param::RetVal  },
        { L"cbVbaObj",  Source::Vba, Param::Objects },
    };
    const wchar_t* const kPauseFull = L"cbPauseFull";

    const Toggle* FindToggle(const wchar_t* id)
    {
        if (!id) return nullptr;
        for (const Toggle& t : kToggles) if (wcscmp(t.id, id) == 0) return &t;
        return nullptr;
    }
    bool Is(const wchar_t* id, const wchar_t* what) { return id && wcscmp(id, what) == 0; }
}

// Three large buttons on the Developer tab; a wrong built-in id drops the lot.
const wchar_t* const kCustomUi =
L"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad' loadImage='OnLoadImage'>"
 L"<ribbon><tabs>"
  L"<tab idMso='TabDeveloper'>"
   L"<group id='grpXRay' label='XRayXL'>"
    L"<button id='btnArm' label='Arm' size='large' getEnabled='GetEnabled' onAction='OnArm'"
           L" imageMso='MacroRecord'"
           L" screentip='Start recording'"
           L" supertip='The same as Application.Run &quot;XRayXL_Arm&quot;.'/>"
    L"<button id='btnDisarm' label='Disarm' size='large' getEnabled='GetEnabled' onAction='OnDisarm'"
           L" image='disarm'"
           L" screentip='Stop, flush and close the trace'"
           L" supertip='The same as Application.Run &quot;XRayXL_Disarm&quot;.'/>"
    L"<button id='btnOptions' label='Options' size='large' getEnabled='GetEnabled' onAction='OnOptions'"
           L" imageMso='ApplicationOptionsDialog'"
           L" screentip='What to record, and where it goes'"
           L" supertip='Everything XRayXL_SetTraceParam can set. Settings are read at the "
           L"next arm, so they cannot be changed while armed.'/>"
   L"</group>"
  L"</tab>"
 L"</tabs></ribbon></customUI>";

Callback CallbackForName(const wchar_t* name)
{
    if (!name) return CbUnknown;
    if (_wcsicmp(name, L"OnRibbonLoad") == 0)         return CbOnLoad;
    if (_wcsicmp(name, L"OnArm") == 0)                return CbOnArm;
    if (_wcsicmp(name, L"OnDisarm") == 0)             return CbOnDisarm;
    if (_wcsicmp(name, L"GetEnabled") == 0)           return CbGetEnabled;
    if (_wcsicmp(name, L"OnOptions") == 0)           return CbOnOptions;
    if (_wcsicmp(name, L"OnLoadImage") == 0)          return CbLoadImage;
    return CbUnknown;
}

bool IsToggle(const wchar_t* id) { return FindToggle(id) != nullptr || Is(id, kPauseFull); }

bool IsDepthControl(const wchar_t* id, Source& src)
{
    if (Is(id, L"ddXllDepth")) { src = Source::Xll; return true; }
    if (Is(id, L"ddVbaDepth")) { src = Source::Vba; return true; }
    return false;
}

bool KnownControl(const wchar_t* id)
{
    Source ignored = Source::Xll;
    return Is(id, L"btnArm") || Is(id, L"btnDisarm") || Is(id, L"btnOptions")
        || IsToggle(id) || IsDepthControl(id, ignored);
}

bool EnabledFor(const wchar_t* id, bool armed)
{
    if (Is(id, L"btnArm"))    return !armed;
    if (Is(id, L"btnDisarm")) return  armed;
    // always reachable: the dialog greys what cannot be changed, which explains itself
    if (Is(id, L"btnOptions")) return true;
    return !armed;              // every setting, refused while armed
}

void BufferText(wchar_t* out, int cap)
{
    if (!out || cap <= 0) return;
    app::params::FormatBufferW(core::modes::GetBufferBytes(), out, cap);
}

bool SetBufferText(const wchar_t* text)
{
    if (!text) return false;
    wchar_t up[24];
    int i = 0;
    for (; text[i] && i < 23; ++i) up[i] = static_cast<wchar_t>(towupper(text[i]));
    up[i] = 0;
    unsigned long long bytes = 0;
    if (!app::params::ParseBufferText(up, bytes)) return false;
    // the command surface's floor: 0 is synchronous, the smallest ring 16KB, nothing between
    if (bytes != 0 && bytes < app::params::kBufMinRing) return false;
    core::modes::SetBufferBytes(static_cast<std::size_t>(bytes));
    return true;
}

bool ReadToggle(const wchar_t* id)
{
    if (Is(id, kPauseFull)) return core::modes::GetPauseOnFull();
    const Toggle* t = FindToggle(id);
    if (!t) return false;
    switch (t->param)
    {
    case Param::Args:    return core::modes::GetArgs(t->src);
    case Param::RetVal:  return core::modes::GetRetVal(t->src);
    case Param::Objects: return core::modes::GetObjects(t->src);
    default:             return false;
    }
}

void WriteToggle(const wchar_t* id, bool on)
{
    if (Is(id, kPauseFull)) { core::modes::SetPauseOnFull(on); return; }
    const Toggle* t = FindToggle(id);
    if (!t) return;
    switch (t->param)
    {
    case Param::Args:    core::modes::SetArgs(t->src, on);    break;
    case Param::RetVal:  core::modes::SetRetVal(t->src, on);  break;
    case Param::Objects: core::modes::SetObjects(t->src, on); break;
    default: break;
    }
}

int DepthIndex(const wchar_t* id)
{
    Source src = Source::Xll;
    if (!IsDepthControl(id, src)) return -1;
    return static_cast<int>(core::modes::GetDepth(src));
}

bool SetDepthIndex(const wchar_t* id, int index)
{
    Source src = Source::Xll;
    if (!IsDepthControl(id, src)) return false;
    // the index is the Depth enum value; anything else changes nothing
    if (index < 0 || index > 2) return false;
    core::modes::SetDepth(src, static_cast<Depth>(index));
    return true;
}
}
}
}
