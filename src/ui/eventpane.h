#pragma once
#include <windows.h>
#include <string>
#include "core/eventlist.h"

// The Events page's list: a heading and a check box per event, in a pane that scrolls. The boxes
// are owner-drawn like the dialog's others, and each tells a screen reader its role and state.
namespace ui
{
namespace eventpane
{
    // The WM_COMMAND code the pane sends its dialog when a tick changes.
    constexpr WORD kChanged = 0x5001;
    // Control ids: a group's heading, and an event's box by catalogue index.
    constexpr int kFirstGroupId = 1900;
    constexpr int kFirstEventId = 2000;

    HWND Create(HWND dlg, int id);
    void Set(HWND pane, core::events::Mask selected, core::events::Mask available);
    core::events::Mask Selected(HWND pane);
    void Enable(HWND pane, bool on);

    // What hovering `id` says: the parameters an event records, or what a heading's next click does.
    std::wstring TipOf(HWND pane, int id);

    // A group's box: 0 clear, 1 all available ticked, 2 mixed.
    int GroupState(HWND pane, int group);

    // The page's summary, kept in step with the ticks: a drop-down of presets, which shows the
    // one the ticks match or Custom, and a label counting the ticked events this Excel has.
    void FillPresets(HWND combo);
    void ShowSummary(HWND pane, HWND combo, HWND count);
    // The drop-down's choice ticks exactly that preset's events among those this Excel has; the
    // rest keep their ticks.
    void ChoosePreset(HWND pane, HWND combo, HWND count);
}
}
