#pragma once

#include <windows.h>
#include <oaidl.h>

namespace ui
{
namespace ribbon
{
namespace art
{
    // Each button's picture, sized for a large button at the DPI of `dpiOf`. Arm, Disarm, Tail and
    // Perfetto are a window holding a list, badged with a red dot, a dark square, a green triangle
    // pointing down, or three nested bars. Options is a checklist; Diagnostics is a folded sheet
    // badged with two red bars.
    IDispatch* ArmPicture(HWND dpiOf);
    IDispatch* DisarmPicture(HWND dpiOf);
    IDispatch* TailPicture(HWND dpiOf);
    IDispatch* PerfettoPicture(HWND dpiOf);
    IDispatch* OptionsPicture(HWND dpiOf);
    IDispatch* DiagnosticsPicture(HWND dpiOf);
}
}
}
