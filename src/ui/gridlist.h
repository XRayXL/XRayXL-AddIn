#pragma once

#include "diag/snapshot.h"

#include <windows.h>
#include <string>
#include <vector>

// A multi-column list drawn like the one in Excel's Name Manager: a white header with a tint on
// the sorted column, no gridlines under it, a full-row selection, and the system's scroll bars.
// Rows are chosen one at a time, by Ctrl, by Shift, or all at once.

namespace ui
{
namespace grid
{
    // The WM_COMMAND notification code the list sends its parent: right-clicked, and the
    // parent owns the menu.
    constexpr WORD kContextMenu = 0x7A02;

    // Registers the window class on first use. `id` is the child id WM_COMMAND carries.
    HWND Create(HWND parent, int id);

    // Drops the window class once the dialog holding the list has closed. Without it an XLL
    // Excel unloads and loads again would meet its own class with a window procedure that is
    // no longer mapped.
    void Unregister();

    // Replaces the contents. The sort, the filter and the column widths are reset with it.
    void SetTable(HWND list, diag::Table table);

    // Keeps only the rows holding `text` in some cell, case-insensitively. Empty shows all.
    void SetFilter(HWND list, const wchar_t* text);

    // The rows as they are shown -- filtered and sorted -- with the columns. Nothing in the
    // dialog needs it; it is how a test sees what the filter and the sort did.
    diag::Table Shown(HWND list);

    // Only the chosen rows, in the order they are shown. Empty when nothing is chosen.
    diag::Table Chosen(HWND list);

    int RowCount(HWND list);        // rows shown, after the filter
    int ChosenCount(HWND list);
    void SelectAll(HWND list);

    // Where the keyboard is, and so which row carries the focus ring. -1 for none.
    int Focused(HWND list);

    // Where the last right-click was, in screen coordinates. From the keyboard, the focus row.
    POINT MenuPoint(HWND list);

    // Column widths and the header are in pixels, so the DPI is pushed in rather than read.
    void SetDpi(HWND list, int dpi);
}
}
