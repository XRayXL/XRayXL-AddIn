#include "core/log.h"
#include "core/xltype.h"
#include "xllregistry.h"
#include "core/excel_api.h"
#include "core/excel_om.h"
#include "core/clock.h"
#include "xlcall.h"

#include <windows.h>
#include <oaidl.h>
#include <sstream>
#include <cstdio>

namespace xll
{
    namespace
    {
        std::wstring BstrToWide(BSTR b)
        {
            if (b == nullptr) return L"";
            return std::wstring(b, SysStringLen(b));
        }

        std::wstring VariantToWide(const VARIANT& v)
        {
            if (v.vt == VT_BSTR) return BstrToWide(v.bstrVal);
            if (v.vt == (VT_BSTR | VT_BYREF) && v.pbstrVal) return BstrToWide(*v.pbstrVal);
            return L"";
        }

        // The two COM lifetimes this file holds. Enumerate has five exits after taking an
        // Application and a result VARIANT, so the compiler holds the releases.
        template <class T>
        class ComPtr
        {
        public:
            explicit ComPtr(T* p) : m_p(p) {}   // ADOPTS an already-AddRefed pointer
            ~ComPtr() { Release(); }
            ComPtr(const ComPtr&) = delete;
            ComPtr& operator=(const ComPtr&) = delete;
            T* operator->() const { return m_p; }
            explicit operator bool() const { return m_p != nullptr; }
            // Explicit, because Enumerate lets the Application go as soon as
            // Invoke returns rather than across the array walk.
            void Release() { if (m_p) { m_p->Release(); m_p = nullptr; } }
        private:
            T* m_p = nullptr;
        };

        class Variant
        {
        public:
            Variant() { VariantInit(&m_v); }
            ~Variant() { VariantClear(&m_v); }
            Variant(const Variant&) = delete;
            Variant& operator=(const Variant&) = delete;
            VARIANT*       Addr()       { return &m_v; }   // to be filled in
            const VARIANT& Get() const  { return m_v; }
            VARTYPE        Type() const { return m_v.vt; }
        private:
            VARIANT m_v;
        };
    }

    bool Enumerate(std::vector<Registration>& out, std::string& log)
    {
        std::ostringstream l;
        ComPtr<IDispatch> app(core::excelom::AcquireApplication(l));
        if (!app) { log = l.str() + "  no Application\n"; return false; }

        DISPID id = 0;
        OLECHAR* nm = const_cast<OLECHAR*>(L"RegisteredFunctions");
        HRESULT hr = app->GetIDsOfNames(IID_NULL, &nm, 1, LOCALE_USER_DEFAULT, &id);
        if (FAILED(hr)) { log = "  no RegisteredFunctions\n"; return false; }

        DISPPARAMS noArgs = { nullptr, nullptr, 0, 0 };
        Variant res;
        hr = app->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_PROPERTYGET | DISPATCH_METHOD, &noArgs, res.Addr(), nullptr, nullptr);
        app.Release();     // not held across the walk below

        if (FAILED(hr)) { log = "  RegisteredFunctions failed\n"; return false; }

        // Null means nothing is registered: an empty table, not an error, and
        // it must not read as a failure to look.
        if (res.Type() == VT_NULL || res.Type() == VT_EMPTY)
        {
            log = "  RegisteredFunctions: none registered\n";
            return true;
        }

        SAFEARRAY* sa = nullptr;
        if (res.Type() == (VT_ARRAY | VT_VARIANT)) sa = res.Get().parray;
        else if (res.Type() == (VT_ARRAY | VT_VARIANT | VT_BYREF) && res.Get().pparray)
            sa = *res.Get().pparray;
        if (sa == nullptr) { log = "  RegisteredFunctions: not an array\n"; return true; }

        LONG r1 = 0, r2 = 0, c1 = 0, c2 = 0;
        SafeArrayGetLBound(sa, 1, &r1); SafeArrayGetUBound(sa, 1, &r2);
        SafeArrayGetLBound(sa, 2, &c1); SafeArrayGetUBound(sa, 2, &c2);

        for (LONG row = r1; row <= r2; row++)
        {
            Registration reg;
            for (LONG col = c1; col <= c2 && col < c1 + 3; col++)
            {
                LONG idx[2] = { row, col };
                Variant cell;      // fresh per cell, cleared on the way out
                if (SUCCEEDED(SafeArrayGetElement(sa, idx, cell.Addr())))
                {
                    const std::wstring s = VariantToWide(cell.Get());
                    if (col == c1)     reg.module = s;
                    else if (col == c1 + 1) reg.procedure = s;
                    else                    reg.typeText = s;
                }
            }
            if (!reg.module.empty() && !reg.procedure.empty()) out.push_back(reg);
        }

        std::ostringstream o;
        o << "  RegisteredFunctions: " << out.size() << " rows\n";
        log = o.str();
        return true;
    }

    // ---- what the name costs (internal) --------------------------------------
    namespace
    {
        ResolveCost g_cost;
        using core::QpcMicros;
    }

    // ---- the real name ------------------------------------------------------
    ResolveCost TakeResolveCost()
    {
        const ResolveCost c = g_cost;
        g_cost = ResolveCost{};
        return c;
    }

    std::wstring ResolveFunctionText(const std::wstring& module,
                                     const std::wstring& procedure,
                                     const std::wstring& typeText)
    {
        core::PascalStr m = core::MakeStr(module.c_str());
        core::PascalStr p = core::MakeStr(procedure.c_str());
        core::PascalStr t = core::MakeStr(typeText.c_str());

        g_cost.calls++;
        const long long t0 = QpcMicros();

        XLOPER12 regId{};
        if (Excel12(xlfRegisterId, &regId, 3, &m.oper, &p.oper, &t.oper) != xlretSuccess)
        {
            g_cost.regIdUs += QpcMicros() - t0;
            return L"";
        }
        double idNum = 0;
        const int rt = regId.xltype & core::kXlTypeMask;
        if (rt == xltypeNum)      idNum = regId.val.num;
        else if (rt == xltypeInt) idNum = regId.val.w;
        else { Excel12(xlFree, nullptr, 1, &regId); g_cost.regIdUs += QpcMicros() - t0; return L""; }
        Excel12(xlFree, nullptr, 1, &regId);
        const long long t1 = QpcMicros();
        g_cost.regIdUs += t1 - t0;



        // xlfGetDef(<register id>, <MISSING>, 3) gives the display name. document_text must be
        // genuinely missing (an empty string fails) and pxTypeNum is 1 or 3, never 2. It
        // resolves names for other add-ins too.
        XLOPER12 idOper{};
        idOper.xltype = xltypeNum;
        idOper.val.num = idNum;
        XLOPER12 missing{};
        missing.xltype = xltypeMissing;
        XLOPER12 typeNum{};
        typeNum.xltype = xltypeInt;
        typeNum.val.w = 3;                      // 1 and 3 both work; 3 is broadest

        XLOPER12 defRes{};
        const int drc = Excel12(xlfGetDef, &defRes, 3, &idOper, &missing, &typeNum);

        std::wstring name;
        if (drc == xlretSuccess && (defRes.xltype & core::kXlTypeMask) == xltypeStr &&
            defRes.val.str != nullptr)
        {
            name.assign(defRes.val.str + 1, static_cast<size_t>(defRes.val.str[0]));
        }
        Excel12(xlFree, nullptr, 1, &defRes);
        g_cost.getDefUs += QpcMicros() - t1;

        // Empty rather than wrong: a registration Excel will not name is
        // reported unnamed, and the caller falls back to the export name.
        if (!name.empty()) g_cost.resolved++;
        return name;
    }
}
