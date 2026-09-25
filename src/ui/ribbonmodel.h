// The ribbon's and the dialog's decisions, free of COM and Excel so a test can reach them.
#pragma once
#include "core/tracemodes.h"

namespace ui
{
namespace ribbon
{
namespace model
{
    // The customUI XML Office asks for through IRibbonExtensibility::GetCustomUI.
    extern const wchar_t* const kCustomUi;

    // The callbacks the XML names; the values are the IDispatch dispids.
    enum Callback
    {
        CbUnknown = 0,
        CbOnLoad = 1, CbOnArm, CbOnDisarm, CbGetEnabled, CbOnOptions, CbOnDiagnostics,
        CbLoadImage, CbOnTail
    };
    Callback CallbackForName(const wchar_t* name);   // case-insensitive, as Office asks

    // False for an id the handlers do not understand: a dead control.
    bool KnownControl(const wchar_t* id);

    // The setters' rule: refused while armed. Says what may be pressed; enforces nothing.
    // `traced`: a trace file has been named this session, so Tail has something to follow.
    bool EnabledFor(const wchar_t* id, bool armed, bool traced);

    // ---- the settings, read and written by the Options dialog; the ids are the setting keys ----

    // Check boxes. ReadToggle is false for an id that is not one.
    bool IsToggle(const wchar_t* id);
    bool ReadToggle(const wchar_t* id);
    void WriteToggle(const wchar_t* id, bool on);

    // Depth drop-downs. The index is the Depth enum value; -1 for an id that is not one.
    bool IsDepthControl(const wchar_t* id, core::modes::Source& src);
    int  DepthIndex(const wchar_t* id);
    bool SetDepthIndex(const wchar_t* id, int index);

    // The ring size in SetTraceParam's words: "64MB", "512KB", "0".
    void BufferText(wchar_t* out, int cap);
    bool SetBufferText(const wchar_t* text);
    bool ParseBufferBox(const wchar_t* text, std::size_t& bytes);   // false: not a valid size
}
}
}
