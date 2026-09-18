#pragma once
#include <windows.h>
#include <oaidl.h>
#include <sstream>

namespace core
{

// The Excel Application IDispatch for THIS process.
namespace excelom
{
    // xlGetHwnd, then the XLMAIN window in our own process, then AccessibleObjectFromWindow on its
    // EXCEL7 child. Returns an AddRef'd pointer the caller must Release, or nullptr.
    IDispatch* AcquireApplication(std::ostringstream& log);

    // Make sure VBA is loaded, by asking Excel for it rather than mapping VBE7 ourselves. Arm path only.
    bool EnsureVbaLoaded(std::ostringstream& log);
}
}   // namespace core
