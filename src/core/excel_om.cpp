// Excel's object model, reached as a COM CLIENT -- arming enumerates Excel's
// registered functions, which means talking to Application.
//
// Client, never server, and the two are not the same risk: a server hands COM a
// pointer into this module and so gives it a second lifetime owner that unloads
// at CoUninitialize whatever DllCanUnloadNow says. A client calls somebody
// else's object and releases it.
#include "excel_om.h"
#include "excel_api.h"
#include "xlcall.h"

#include <windows.h>
#include <oaidl.h>
#include <oleacc.h>

namespace core
{

namespace
{
    // IDispatch property-get returning an object (IDispatch*), AddRef'd.
    IDispatch* GetDispProp(IDispatch* obj, const wchar_t* name)
    {
        if (!obj) return nullptr;
        DISPID id = 0;
        OLECHAR* n = const_cast<OLECHAR*>(name);
        if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return nullptr;
        DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
        VARIANT out; VariantInit(&out);
        HRESULT hr = obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &dp, &out, nullptr, nullptr);
        IDispatch* r = nullptr;
        if (SUCCEEDED(hr) && out.vt == VT_DISPATCH && out.pdispVal) { r = out.pdispVal; r->AddRef(); }
        VariantClear(&out);
        return r;
    }

    struct FindChild { HWND found; };

    BOOL CALLBACK LookForExcel7(HWND hwnd, LPARAM param)
    {
        wchar_t cls[64] = { 0 };
        GetClassNameW(hwnd, cls, 63);
        if (wcscmp(cls, L"EXCEL7") == 0)
        {
            reinterpret_cast<FindChild*>(param)->found = hwnd;
            return FALSE;
        }
        return TRUE;
    }
}

namespace excelom
{
    IDispatch* AcquireApplication(std::ostringstream& log)
    {
        XLOPER12 hres{};
        if (Excel12(xlGetHwnd, &hres, 0) != xlretSuccess)
        {
            log << "  xlGetHwnd failed" << std::endl;
            return nullptr;
        }

        // xlGetHwnd returns an int -- on x64 the low half of a real HWND, so
        // it is recovered by matching. FindWindowExW enumerates the whole
        // desktop and a working machine often runs several Excels, so the
        // process filter is what stops us binding to somebody else's instance.
        HWND top = nullptr;
        {
            int wanted = hres.val.w;
            DWORD mine = GetCurrentProcessId();
            HWND h = nullptr;
            while ((h = FindWindowExW(nullptr, h, L"XLMAIN", nullptr)) != nullptr)
            {
                DWORD owner = 0;
                GetWindowThreadProcessId(h, &owner);
                if (owner != mine) continue;
                if ((static_cast<int>(reinterpret_cast<INT_PTR>(h)) & 0xFFFF) == (wanted & 0xFFFF))
                { top = h; break; }
                if (top == nullptr) top = h;    // ours, but not the one named
            }
        }
        Excel12(xlFree, nullptr, 1, &hres);
        if (top == nullptr) { log << "  no XLMAIN window" << std::endl; return nullptr; }

        FindChild fc{ nullptr };
        EnumChildWindows(top, LookForExcel7, reinterpret_cast<LPARAM>(&fc));
        if (fc.found == nullptr) { log << "  no EXCEL7 child window" << std::endl; return nullptr; }

        IDispatch* window = nullptr;
        HRESULT hr = AccessibleObjectFromWindow(fc.found, static_cast<DWORD>(OBJID_NATIVEOM),
                                                IID_IDispatch, reinterpret_cast<void**>(&window));
        if (FAILED(hr) || window == nullptr)
        {
            log << "  AccessibleObjectFromWindow failed 0x" << std::hex << hr << std::dec << std::endl;
            return nullptr;
        }

        IDispatch* app = GetDispProp(window, L"Application");
        window->Release();
        if (app == nullptr) log << "  could not reach Application" << std::endl;
        return app;
    }
}
}   // namespace core
