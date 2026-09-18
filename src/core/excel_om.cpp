// Excel's object model, reached as a COM client: arming enumerates Excel's registered functions
// through Application.
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

        // xlGetHwnd returns only the low half of the HWND on x64, so it is matched against the
        // windows of our own process; a machine often runs several Excels.
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
    // Reading ActiveWorkbook.VBProject makes Excel load and initialise VBA itself. Not
    // Application.VBE, which is refused before anything loads when project trust is off.
    bool EnsureVbaLoaded(std::ostringstream& log)
    {
        if (GetModuleHandleW(L"VBE7.DLL") != nullptr) return true;   // Excel already did

        IDispatch* app = AcquireApplication(log);
        if (!app) { log << "  VBA is not loaded and the object model is not reachable to ask" << std::endl; return false; }

        IDispatch* wb = GetDispProp(app, L"ActiveWorkbook");
        if (wb)
        {
            // The value is discarded and a refusal is fine: asking is what loads it.
            IDispatch* prj = GetDispProp(wb, L"VBProject");
            if (prj) prj->Release();
            wb->Release();
        }
        else log << "  no ActiveWorkbook to ask for a VBProject" << std::endl;
        app->Release();

        const bool loaded = GetModuleHandleW(L"VBE7.DLL") != nullptr;
        log << "  VBA was not loaded; asked Excel for it -> "
            << (loaded ? "loaded" : "still absent") << std::endl;
        return loaded;
    }

}
}   // namespace core
