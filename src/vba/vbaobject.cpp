#include "vbaobject.h"
#include "vbaretdecode.h"
#include <windows.h>
#include <oaidl.h>
#include <oleauto.h>
#include <ocidl.h>
#include <cstdio>
#include <cstring>

namespace vba
{
namespace
{
    // Published IIDs. A wrong GUID never matches; the object is then named from its type info.
    const GUID kIidRange     = { 0x00020846, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };
    const GUID kIidWorksheet = { 0x000208D8, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };
    const GUID kIidWorkbook  = { 0x000208DA, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };

    // So a class that stops being recognised is visible rather than silently absent.
    volatile LONG64 g_described = 0, g_namedOnly = 0, g_unknown = 0;

    // Interfaces and a VARIANT held inside the guarded call, so a fault part-way
    // through can still release them. Trivially destructible, so SEH allows it.
    struct ComHold
    {
        IUnknown* p[8];
        int       n;
        VARIANT*  var;
        void Add(IUnknown* u)  { if (u && n < 8) p[n++] = u; }
        void Drop(IUnknown* u) { for (int i = n - 1; i >= 0; --i) if (p[i] == u) { p[i] = p[--n]; return; } }
        void ReleaseAll()
        {
            while (n > 0) { IUnknown* u = p[--n]; if (u) u->Release(); }
            if (var) { VariantClear(var); var = nullptr; }
        }
    };
    __declspec(thread) ComHold* t_hold = nullptr;
    void Hold(IUnknown* u)   { if (t_hold) t_hold->Add(u); }
    void Unhold(IUnknown* u) { if (t_hold) t_hold->Drop(u); }

    // Value2 of `A:A` would materialise a 25 MB VARIANT array inside a calculation.
    constexpr long kMaxCellsToRead = 4096;

    bool Answers(IDispatch* d, const GUID& iid)
    {
        if (!d) return false;
        void* p = nullptr;
        if (FAILED(d->QueryInterface(iid, &p)) || !p) return false;
        reinterpret_cast<IUnknown*>(p)->Release();
        return true;
    }

    // `args` in declaration order; DISPPARAMS wants them reversed.
    bool GetProp(IDispatch* d, const wchar_t* name, VARIANT& out, VARIANT* args, int argc)
    {
        VariantInit(&out);
        if (!d) return false;
        DISPID id = 0;
        OLECHAR* n = const_cast<OLECHAR*>(name);
        if (FAILED(d->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return false;

        VARIANT rev[8];
        for (int i = 0; i < argc && i < 8; ++i) rev[i] = args[argc - 1 - i];
        DISPPARAMS dp = { argc ? rev : nullptr, nullptr,
                          static_cast<UINT>(argc), 0 };
        return SUCCEEDED(d->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &dp, &out, nullptr, nullptr));
    }

    bool GetStr(IDispatch* d, const wchar_t* name, char* out, int cap,
                VARIANT* args = nullptr, int argc = 0)
    {
        VARIANT v;
        if (!GetProp(d, name, v, args, argc)) return false;
        bool ok = false;
        if (v.vt == VT_BSTR && v.bstrVal)
        {
            const int n = WideCharToMultiByte(CP_UTF8, 0, v.bstrVal, -1, out, cap, nullptr, nullptr);
            ok = (n > 0);
        }
        VariantClear(&v);
        return ok;
    }

    bool GetLong(IDispatch* d, const wchar_t* name, long& out)
    {
        VARIANT v;
        if (!GetProp(d, name, v, nullptr, 0)) return false;
        bool ok = false;
        if (v.vt == VT_I4)      { out = v.lVal;                      ok = true; }
        else if (v.vt == VT_R8) { out = static_cast<long>(v.dblVal); ok = true; }
        VariantClear(&v);
        return ok;
    }

    IDispatch* GetObj(IDispatch* d, const wchar_t* name)
    {
        VARIANT v;
        if (!GetProp(d, name, v, nullptr, 0)) return nullptr;
        IDispatch* r = nullptr;
        if (v.vt == VT_DISPATCH && v.pdispVal) { r = v.pdispVal; r->AddRef(); }
        VariantClear(&v);
        return r;
    }

    // What VBA's TypeName() reads, passed through as given, underscore included (`_Worksheet`).
    bool NameOf(ITypeInfo* ti, char* out, int cap)
    {
        BSTR name = nullptr;
        if (FAILED(ti->GetDocumentation(MEMBERID_NIL, &name, nullptr, nullptr, nullptr)) || !name)
            return false;
        const bool ok = (WideCharToMultiByte(CP_UTF8, 0, name, -1, out, cap, nullptr, nullptr) > 0);
        SysFreeString(name);
        return ok;
    }

    bool TypeName(IDispatch* d, char* out, int cap)
    {
        // The coclass first: GetTypeInfo answers `_Collection` where TypeName() says `Collection`.
        // Excel's sheet object gives no coclass.
        IProvideClassInfo* pci = nullptr;
        if (SUCCEEDED(d->QueryInterface(IID_IProvideClassInfo,
                                        reinterpret_cast<void**>(&pci))) && pci)
        {
            Hold(pci);
            ITypeInfo* ci = nullptr;
            const bool got = SUCCEEDED(pci->GetClassInfo(&ci)) && ci;
            bool ok = false;
            if (got) { Hold(ci); ok = NameOf(ci, out, cap); Unhold(ci); ci->Release(); }
            Unhold(pci); pci->Release();
            if (ok) return true;
        }

        UINT n = 0;
        if (FAILED(d->GetTypeInfoCount(&n)) || n == 0) return false;
        ITypeInfo* ti = nullptr;
        if (FAILED(d->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ti)) || !ti) return false;
        Hold(ti);
        const bool ok = NameOf(ti, out, cap);
        Unhold(ti); ti->Release();
        return ok;
    }

    // `[Book1]Sheet1!A1:B2`, the form `callerref` uses, so an object argument and a calling
    // cell read alike.
    bool RangeAddress(IDispatch* r, char* out, int cap)
    {
        VARIANT a[4];
        for (int i = 0; i < 4; ++i) VariantInit(&a[i]);
        a[0].vt = VT_BOOL; a[0].boolVal = VARIANT_FALSE;   // RowAbsolute
        a[1].vt = VT_BOOL; a[1].boolVal = VARIANT_FALSE;   // ColumnAbsolute
        a[2].vt = VT_I4;   a[2].lVal    = 1;               // xlA1
        a[3].vt = VT_BOOL; a[3].boolVal = VARIANT_TRUE;    // External
        return GetStr(r, L"Address", out, cap, a, 4);
    }

    // Contents are declined, not cut, when too many or when the count is unknown: an unknown
    // count may not be assumed small.
    bool DescribeRange(IDispatch* r, std::uint64_t ptr, core::ValueWriter& w)
    {
        char addr[512];
        if (!RangeAddress(r, addr, sizeof(addr))) return false;
        w.BeginObject("Range", ptr, addr);

        long cells = 0;
        const bool haveCount = GetLong(r, L"Count", cells);
        VARIANT v;
        if (haveCount && cells <= kMaxCellsToRead && GetProp(r, L"Value2", v, nullptr, 0))
        {
            if (t_hold) t_hold->var = &v;
            DescribeVariantValue(reinterpret_cast<std::uint64_t>(&v), w);
            if (t_hold) t_hold->var = nullptr;
            VariantClear(&v);
        }
        w.EndObject();
        return true;
    }

    bool DescribeWorksheet(IDispatch* ws, std::uint64_t ptr, core::ValueWriter& w)
    {
        char sheet[256];
        if (!GetStr(ws, L"Name", sheet, sizeof(sheet))) return false;
        char book[256]; book[0] = 0;
        if (IDispatch* parent = GetObj(ws, L"Parent"))
        {
            Hold(parent);
            // GetStr can leave the buffer unterminated on failure.
            if (!GetStr(parent, L"Name", book, sizeof(book))) book[0] = 0;
            Unhold(parent); parent->Release();
        }
        // `[Book1]Sheet1`, as callerref; without the book, the sheet alone rather than a guess.
        char where[600];
        if (book[0]) _snprintf_s(where, _TRUNCATE, "[%s]%s", book, sheet);
        else         _snprintf_s(where, _TRUNCATE, "%s", sheet);
        w.BeginObject("Worksheet", ptr, where);
        w.EndObject();
        return true;
    }

    bool DescribeWorkbook(IDispatch* wb, std::uint64_t ptr, core::ValueWriter& w)
    {
        char book[256];
        if (!GetStr(wb, L"Name", book, sizeof(book))) return false;
        char where[300];
        _snprintf_s(where, _TRUNCATE, "[%s]", book);
        w.BeginObject("Workbook", ptr, where);
        w.EndObject();
        return true;
    }

    // No object that needs unwinding, so the caller can wrap it in SEH.
    bool DescribeInner(IDispatch* d, std::uint64_t ptr, core::ValueWriter& w)
    {
        // Named by the QueryInterface match: the type info would say `_Worksheet`.
        const char* known = nullptr;
        bool detailed = false;
        if (Answers(d, kIidRange))
        {
            known = "Range";
            detailed = DescribeRange(d, ptr, w);
        }
        else if (Answers(d, kIidWorksheet))
        {
            known = "Worksheet";
            detailed = DescribeWorksheet(d, ptr, w);
        }
        else if (Answers(d, kIidWorkbook))
        {
            known = "Workbook";
            detailed = DescribeWorkbook(d, ptr, w);
        }

        if (detailed) { InterlockedIncrement64(&g_described); return true; }

        char cls[128];
        if (known) strncpy_s(cls, known, _TRUNCATE);
        else if (!TypeName(d, cls, sizeof(cls)) || !cls[0])
        {
            InterlockedIncrement64(&g_unknown);
            return false;
        }
        w.BeginObject(cls, ptr, nullptr);
        w.EndObject();
        InterlockedIncrement64(&g_namedOnly);
        return true;
    }
}

void ObjectTotals(long long& described, long long& namedOnly, long long& unknown)
{
    described = InterlockedCompareExchange64(&g_described, 0, 0);
    namedOnly = InterlockedCompareExchange64(&g_namedOnly, 0, 0);
    unknown   = InterlockedCompareExchange64(&g_unknown,   0, 0);
}

void ResetObjectTotals()
{
    InterlockedExchange64(&g_described, 0);
    InterlockedExchange64(&g_namedOnly, 0);
    InterlockedExchange64(&g_unknown,   0);
}


bool DescribeObjectDetail(std::uint64_t ptr, core::ValueWriter& w)
{
    if (!ptr) return false;
    const core::ValueWriter::Mark mark = w.Save();
    ComHold hold;
    hold.n = 0; hold.var = nullptr;
    t_hold = &hold;
    __try
    {
        // Callers pass VT_UNKNOWN too; QueryInterface is the only safe call on a bare IUnknown.
        IDispatch* d = nullptr;
        IUnknown*  unk = reinterpret_cast<IUnknown*>(ptr);
        if (FAILED(unk->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&d))) || !d)
        {
            t_hold = nullptr;
            return false;
        }
        Hold(d);
        const bool ok = DescribeInner(d, ptr, w);
        Unhold(d); d->Release();
        t_hold = nullptr;
        return ok;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hold.ReleaseAll();
        t_hold = nullptr;
        w.Restore(mark);
        return false;
    }
}
}   // namespace vba
