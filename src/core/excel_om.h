#pragma once
#include <windows.h>
#include <oaidl.h>
#include <sstream>

namespace core
{

// The Excel Application IDispatch for THIS process.
namespace excelom
{
    // xlGetHwnd -> the XLMAIN window carrying that low half IN OUR OWN PROCESS
    // (a machine often runs several Excels, so the process filter is not optional) ->
    // AccessibleObjectFromWindow on its EXCEL7 child -> .Application.
    //
    // Returns an AddRef'd pointer the caller must Release, or nullptr.
    IDispatch* AcquireApplication(std::ostringstream& log);
}
}   // namespace core
