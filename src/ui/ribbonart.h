#pragma once

#include <windows.h>
#include <oaidl.h>

namespace ui
{
namespace ribbon
{
namespace art
{
    // Office's MacroRecord page with its badge redrawn 20% bigger around a mark: Office's red dot,
    // a dark stop square, or a green triangle pointing down for following a file as it grows.
    IDispatch* ArmPicture(IDispatch* application, HWND dpiOf);
    IDispatch* DisarmPicture(IDispatch* application, HWND dpiOf);
    IDispatch* TailPicture(IDispatch* application, HWND dpiOf);
}
}
}
