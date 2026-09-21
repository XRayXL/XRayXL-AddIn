#pragma once
// Control ids for the Diagnostics dialog template in XRayXL.rc, shared with diagnosticsdlg.cpp.

#define IDD_XRAY_DIAG      102

#define IDC_DIAG_CATEGORIES 2000
#define IDC_DIAG_PAGEICON   2001
#define IDC_DIAG_HDR        2002
#define IDC_DIAG_SEARCHLBL  2003
#define IDC_DIAG_SEARCH     2004

// The list has no template entry: it is our own window class, created at WM_INITDIALOG.
#define IDC_DIAG_LIST       2013

// The list's context menu.
#define IDM_DIAG_SELECTALL  2100
#define IDM_DIAG_REVEAL     2101
#define IDM_DIAG_COPY       2102
#define IDM_DIAG_EXPORT     2103
