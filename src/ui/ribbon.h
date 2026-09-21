// The ribbon buttons: the only COM object this add-in owns, and the only part allowed to fail.
#pragma once
#include <windows.h>

namespace ui
{
namespace ribbon
{
    // Never throws: on failure it logs and returns. XRAYXL_RIBBON=0 skips it.
    void Start();

    // Idempotent, and waits for nothing.
    void Stop();

    // True while Excel holds the add-in, whose OnDisconnection then reports Excel's exit
    // after the user can no longer cancel it.
    bool Connected();

    // From the exported DllGetClassObject; this add-in's CLSID only.
    HRESULT GetClassObject(REFCLSID rclsid, REFIID riid, void** ppv);
}
}
