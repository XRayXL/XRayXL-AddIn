#include "eventlist.h"
#include <windows.h>
#include <cstdio>
#include <cstring>
#include <initializer_list>

namespace core
{
namespace events
{
    namespace
    {
        using G = Group;
        // DISPIDs from the AppEvents dispinterface in EXCEL.EXE's type library. Order is the
        // dialog's: by group, then as listed.
        const Event kEvents[] = {
            { "SheetChange",                      0x61C, G::CalcData,       true,  false, false, false, "a cell was edited" },
            { "SheetCalculate",                   0x61B, G::CalcData,       true,  false, false, false, "a sheet recalculated" },
            { "AfterCalculate",                   0xA34, G::CalcData,       true,  false, false, false, "all calculation done" },
            { "SheetTableUpdate",                 0xC04, G::CalcData,       true,  false, false, false, "a table was refreshed" },
            { "WorkbookModelChange",              0xC08, G::CalcData,       true,  false, false, false, "the data model changed" },

            { "SheetSelectionChange",             0x616, G::Selection,      false, true,  true,  false, "the selection moved" },
            { "SheetActivate",                    0x619, G::Selection,      false, true,  false, false, "a sheet was activated" },
            { "SheetDeactivate",                  0x61A, G::Selection,      false, true,  false, false, "a sheet was deactivated" },
            { "WorkbookActivate",                 0x620, G::Selection,      false, true,  false, false, "a workbook was activated" },
            { "WorkbookDeactivate",               0x621, G::Selection,      false, true,  false, false, "a workbook was deactivated" },
            { "SheetBeforeDoubleClick",           0x617, G::Selection,      false, true,  false, false, "a cell was double-clicked" },
            { "SheetBeforeRightClick",            0x618, G::Selection,      false, true,  false, false, "a cell was right-clicked" },
            { "SheetFollowHyperlink",             0x73E, G::Selection,      false, true,  false, false, "a hyperlink was followed" },

            { "NewWorkbook",                      0x61D, G::Workbooks,      false, false, false, false, "a workbook was created" },
            { "WorkbookOpen",                     0x61F, G::Workbooks,      false, false, false, false, "a workbook was opened" },
            { "WorkbookBeforeClose",              0x622, G::Workbooks,      false, false, false, false, "a workbook is about to close" },
            { "WorkbookBeforeSave",               0x623, G::Workbooks,      false, false, false, false, "a workbook is about to be saved" },
            { "WorkbookAfterSave",                0xB5F, G::Workbooks,      false, false, false, false, "a workbook was saved" },
            { "WorkbookBeforePrint",              0x624, G::Workbooks,      false, false, false, false, "a workbook is about to print" },
            { "WorkbookNewSheet",                 0x625, G::Workbooks,      false, false, false, false, "a sheet was added" },
            { "WorkbookNewChart",                 0xB60, G::Workbooks,      false, false, false, false, "a chart was added" },
            { "SheetBeforeDelete",                0xC07, G::Workbooks,      false, false, false, false, "a sheet is about to be deleted" },
            { "WorkbookAddinInstall",             0x626, G::Workbooks,      false, false, false, false, "a workbook was installed as an add-in" },
            { "WorkbookAddinUninstall",           0x627, G::Workbooks,      false, false, false, false, "an add-in workbook was uninstalled" },
            { "SheetLensGalleryRenderComplete",   0xC03, G::Workbooks,      false, false, false, false, "a Quick Analysis gallery finished drawing" },

            { "WindowActivate",                   0x614, G::Windows,        false, true,  false, false, "a window was activated" },
            { "WindowDeactivate",                 0x615, G::Windows,        false, true,  false, false, "a window was deactivated" },
            { "WindowResize",                     0x612, G::Windows,        false, false, true,  false, "a window was resized" },

            { "SheetPivotTableUpdate",            0x86D, G::PivotTables,    true,  false, false, false, "a PivotTable was updated" },
            { "SheetPivotTableAfterValueChange",  0xB4F, G::PivotTables,    true,  false, false, false, "PivotTable values were changed" },
            { "SheetPivotTableBeforeAllocateChanges", 0xB50, G::PivotTables, false, false, false, false, "PivotTable changes about to be allocated" },
            { "SheetPivotTableBeforeCommitChanges",   0xB51, G::PivotTables, false, false, false, false, "PivotTable changes about to be committed" },
            { "SheetPivotTableBeforeDiscardChanges",  0xB52, G::PivotTables, false, false, false, false, "PivotTable changes about to be discarded" },
            { "WorkbookPivotTableOpenConnection", 0x871, G::PivotTables,    false, false, false, false, "a PivotTable connection opened" },
            { "WorkbookPivotTableCloseConnection",0x870, G::PivotTables,    false, false, false, false, "a PivotTable connection closed" },

            { "WorkbookRowsetComplete",           0xA33, G::XmlConnections, false, false, false, false, "a drill-through record set finished" },
            { "WorkbookBeforeXmlImport",          0x8F2, G::XmlConnections, false, false, false, false, "XML is about to be imported" },
            { "WorkbookAfterXmlImport",           0x8F3, G::XmlConnections, false, false, false, false, "XML was imported" },
            { "WorkbookBeforeXmlExport",          0x8F4, G::XmlConnections, false, false, false, false, "XML is about to be exported" },
            { "WorkbookAfterXmlExport",           0x8F5, G::XmlConnections, false, false, false, false, "XML was exported" },
            { "WorkbookSync",                     0x8F1, G::XmlConnections, false, false, false, false, "a shared-workspace sync happened" },

            { "ProtectedViewWindowOpen",          0xB57, G::ProtectedView,  false, false, false, false, "a Protected View window opened" },
            { "ProtectedViewWindowActivate",      0xB5D, G::ProtectedView,  false, true,  false, false, "a Protected View window was activated" },
            { "ProtectedViewWindowDeactivate",    0xB5E, G::ProtectedView,  false, true,  false, false, "a Protected View window was deactivated" },
            { "ProtectedViewWindowBeforeEdit",    0xB59, G::ProtectedView,  false, false, false, false, "editing is about to be enabled" },
            { "ProtectedViewWindowBeforeClose",   0xB5A, G::ProtectedView,  false, false, false, false, "a Protected View window is about to close" },
            { "ProtectedViewWindowResize",        0xB5C, G::ProtectedView,  false, false, false, false, "a Protected View window was resized" },

            { "WorkbookBeforeRemoteChange",       0xD16, G::CoAuthoring,    false, false, false, false, "another author's change is about to apply" },
            { "WorkbookAfterRemoteChange",        0xD17, G::CoAuthoring,    true,  false, false, false, "another author's change was applied" },
            { "RemoteSheetChange",                0xD10, G::CoAuthoring,    true,  false, false, true,  "another author edited a cell" },
            { "RemoteWorkbookNewSheet",           0xD18, G::CoAuthoring,    false, false, false, true,  "another author added a sheet" },
            { "RemoteWorkbookNewChart",           0xD19, G::CoAuthoring,    false, false, false, true,  "another author added a chart" },
            { "RemoteSheetBeforeDelete",          0xD13, G::CoAuthoring,    false, false, false, true,  "another author is deleting a sheet" },
            { "RemoteSheetPivotTableUpdate",      0xD14, G::CoAuthoring,    false, false, false, true,  "another author updated a PivotTable" },
        };
        constexpr int kCount = static_cast<int>(sizeof(kEvents) / sizeof(kEvents[0]));
        static_assert(kCount < 64, "the selection is a 64-bit mask, and all 64 set is the unset sentinel");

        Mask Collect(bool Event::*member)
        {
            Mask m = 0;
            for (int i = 0; i < kCount; ++i) if (kEvents[i].*member) m |= Mask(1) << i;
            return m;
        }

        volatile LONG64 g_selected = -1;   // unset: the Calc preset
    }

    int          Count()     { return kCount; }
    const Event& At(int i)   { return kEvents[i]; }

    int Find(const char* name)
    {
        if (!name) return -1;
        for (int i = 0; i < kCount; ++i) if (_stricmp(kEvents[i].name, name) == 0) return i;
        return -1;
    }

    int FindDispid(long dispid)
    {
        for (int i = 0; i < kCount; ++i) if (kEvents[i].dispid == dispid) return i;
        return -1;
    }

    const char* GroupName(Group g)
    {
        switch (g)
        {
        case Group::CalcData:       return "Calculation and data";
        case Group::Selection:      return "Selection and navigation";
        case Group::Workbooks:      return "Workbooks and sheets";
        case Group::Windows:        return "Windows";
        case Group::PivotTables:    return "PivotTables";
        case Group::XmlConnections: return "XML and connections";
        case Group::ProtectedView:  return "Protected View";
        case Group::CoAuthoring:    return "Co-authoring";
        default:                    return "?";
        }
    }

    Mask AllMask() { return kCount == 64 ? ~Mask(0) : (Mask(1) << kCount) - 1; }

    Mask PresetMask(Preset p)
    {
        switch (p)
        {
        case Preset::Calc:      return Collect(&Event::calc);
        case Preset::Selection: return Collect(&Event::selection);
        case Preset::CalcAndSelection: return Collect(&Event::calc) | Collect(&Event::selection);
        case Preset::All:       return AllMask();
        default:                return 0;
        }
    }

    const char* PresetName(Preset p)
    {
        switch (p)
        {
        case Preset::None:      return "None";
        case Preset::Calc:      return "Calc";
        case Preset::Selection: return "Selection";
        case Preset::CalcAndSelection: return "Calc & Selection";
        case Preset::All:       return "All";
        default:                return "Custom";
        }
    }

    Preset PresetOf(Mask m, Mask available)
    {
        for (Preset p : { Preset::None, Preset::Calc, Preset::Selection, Preset::CalcAndSelection, Preset::All })
            if ((PresetMask(p) & available) == (m & available)) return p;
        return Preset::Custom;
    }

    Mask GetSelected()
    {
        const LONG64 v = InterlockedCompareExchange64(&g_selected, 0, 0);
        return v == -1 ? PresetMask(Preset::Calc) : static_cast<Mask>(v);
    }

    void SetSelected(Mask m) { InterlockedExchange64(&g_selected, static_cast<LONG64>(m & AllMask())); }

    void DescribeSelection(Mask m, Mask available, char* out, int cap)
    {
        const Preset p = PresetOf(m, available);
        if (p != Preset::Custom) { _snprintf_s(out, cap, _TRUNCATE, "%s", PresetName(p)); return; }
        int j = 0;
        out[0] = 0;
        for (int i = 0; i < kCount && j < cap - 1; ++i)
        {
            if (!(m & (Mask(1) << i))) continue;
            const int w = _snprintf_s(out + j, cap - j, _TRUNCATE, "%s%s", j ? "," : "", kEvents[i].name);
            if (w < 0) break;
            j += w;
        }
    }
}
}
