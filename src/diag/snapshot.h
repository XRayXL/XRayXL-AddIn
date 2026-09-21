#pragma once

#include <string>
#include <vector>

// What the Diagnostics dialog shows: three tables read from this process, each a header row of
// column names and the rows under it. Nothing here touches Excel, so a test can read them.

namespace diag
{
    struct Column
    {
        std::wstring title;
        int          width96;    // the column's width in pixels at 96 DPI
        bool         rightAlign; // numbers read better against their own edge
    };

    struct Table
    {
        std::vector<Column>                    columns;
        std::vector<std::vector<std::wstring>> rows;   // each as wide as `columns`
    };

    // Every module mapped into this process: name, version, date modified, size and path.
    // The first row is the process image, which for a loaded add-in is Excel.
    Table Modules();

    // The process environment block, one row per variable, sorted by name.
    Table Environment();

    // Memory, handles, GDI and USER objects, threads, CPU time, the OS and the host, as
    // name-and-value rows. The session comes from the caller, so this needs none of it.
    struct Session
    {
        const char*  version   = "";
        bool         armed     = false;
        std::wstring traceFile;
        std::wstring logFile;
    };
    Table Process(const Session& session);
}
