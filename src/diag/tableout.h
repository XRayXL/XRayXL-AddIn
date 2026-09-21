#pragma once

#include "snapshot.h"
#include <string>

// A diagnostics table on its way out: to a file as CSV, or to the clipboard as the
// tab-separated text a spreadsheet pastes into cells.

namespace diag
{
    // UTF-8 with a BOM, so Excel opens a path with accents in it correctly.
    std::string AsCsv(const Table& t);

    std::wstring AsTabbed(const Table& t);

    // Writes `bytes` to `path`, replacing what is there. False if it could not be written.
    bool WriteBytes(const std::wstring& path, const std::string& bytes);

    // %TEMP%\XRayXL\Diagnostics\XRayXL_<page>_<pid>_<stamp>.csv. Empty if the folder refused.
    std::wstring ExportPath(const wchar_t* pageName);
}
