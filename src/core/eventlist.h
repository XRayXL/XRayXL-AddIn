#pragma once
#include <cstdint>

// The events Excel raises on Application (the AppEvents dispinterface), and which of them a trace
// records. An event is known by its DISPID, which never changes between versions, so an older Excel
// simply never raises the ones it lacks.
namespace core
{
namespace events
{
    enum class Group { CalcData, Selection, Workbooks, Windows, PivotTables, XmlConnections,
                       ProtectedView, CoAuthoring, Count_ };
    enum class Preset { None, Calc, Selection, CalcAndSelection, All, Custom };

    struct Event
    {
        const char* name;       // as Excel names it, and as the trace's `function` column reads
        long        dispid;
        Group       group;
        bool        calc;       // in the Calc preset
        bool        selection;  // in the Selection preset
        bool        often;      // fires often enough to swamp a trace
        bool        hidden;     // hidden in the type library (co-authoring)
        const char* shownAs;    // a few plain words for the dialog
    };

    // A bit per catalogue entry, in catalogue order.
    using Mask = std::uint64_t;

    int          Count();
    const Event& At(int i);
    int          Find(const char* name);    // case-insensitive; -1 when unknown
    int          FindDispid(long dispid);   // -1 for an event this catalogue does not know
    const char*  GroupName(Group g);

    Mask        AllMask();
    Mask        PresetMask(Preset p);        // All is AllMask(); Custom and None are 0
    const char* PresetName(Preset p);
    // The preset whose events, among those available, are exactly `m`'s; otherwise Custom.
    Preset      PresetOf(Mask m, Mask available);

    // The recorded set, read once at arm. Default: the Calc preset.
    Mask GetSelected();
    void SetSelected(Mask m);

    // "SheetChange,SheetCalculate", or the preset's name when it is exactly one.
    void DescribeSelection(Mask m, Mask available, char* out, int cap);
}
}
