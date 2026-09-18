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
    // The three classes we detail, by their published interface IDs. QueryInterface is the
    // validation: a wrong GUID never matches, and the object falls through to being named from
    // its type info.
    const GUID kIidRange     = { 0x00020846, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };
    const GUID kIidWorksheet = { 0x000208D8, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };
    const GUID kIidWorkbook  = { 0x000208DA, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };

    // WHAT ACTUALLY HAPPENED, so a class that stops being recognised is
    // visible rather than silently absent. Counted, never gating.
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

    // A RANGE THIS BIG IS NOT READ. `A:A` is 1,048,576 cells and asking Excel
    // for its Value2 would materialise a VARIANT array of about 25 MB inside a
    // calculation -- to render the first 64 of them. The address is still
    // reported; only the contents are declined, and the row shows the difference.
    constexpr long kMaxCellsToRead = 4096;

    bool Answers(IDispatch* d, const GUID& iid)
    {
        if (!d) return false;
        void* p = nullptr;
        if (FAILED(d->QueryInterface(iid, &p)) || !p) return false;
        reinterpret_cast<IUnknown*>(p)->Release();
        return true;
    }

    // A property or method returning a VARIANT. `argc` positional arguments, in
    // DECLARATION order -- reversed here, which is the order DISPPARAMS wants
    // and the single most common way to get this subtly wrong.
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

    // A BSTR property, narrowed. False when it is not a string -- never a
    // partial or a placeholder.
    bool GetStr(IDispatch* d, const wchar_t* name, char* out, int cap,
                VARIANT* args = nullptr, int argc = 0)
    {
        VARIANT v;
        if (!GetProp(d, name, v, args, argc)) return false;
        bool ok = false;
        if (v.vt == VT_BSTR && v.bstrVal)
        {
            const int n = WideCharToMultiByte(CP_ACP, 0, v.bstrVal, -1, out, cap, nullptr, nullptr);
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

    // What the type library calls it, which is what VBA's TypeName() reads. Passed through as
    // given: Excel's own interfaces are `_Worksheet` and `_Workbook`, with the underscore.
    bool NameOf(ITypeInfo* ti, char* out, int cap)
    {
        BSTR name = nullptr;
        if (FAILED(ti->GetDocumentation(MEMBERID_NIL, &name, nullptr, nullptr, nullptr)) || !name)
            return false;
        const bool ok = (WideCharToMultiByte(CP_ACP, 0, name, -1, out, cap, nullptr, nullptr) > 0);
        SysFreeString(name);
        return ok;
    }

    bool TypeName(IDispatch* d, char* out, int cap)
    {
        // The class, not its default interface: IDispatch::GetTypeInfo answers `_Collection`
        // where TypeName() says `Collection`, so IProvideClassInfo is asked for the coclass. It
        // does not always answer; Excel's sheet object gives nothing and falls through to
        // `_Worksheet`.
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

        // No class info: the interface name is what there is, and is reported as
        // it comes -- including any leading underscore, which is a fact about
        // the interface rather than a blemish to tidy.
        UINT n = 0;
        if (FAILED(d->GetTypeInfoCount(&n)) || n == 0) return false;
        ITypeInfo* ti = nullptr;
        if (FAILED(d->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ti)) || !ti) return false;
        Hold(ti);
        const bool ok = NameOf(ti, out, cap);
        Unhold(ti); ti->Release();
        return ok;
    }

    // `Address(RowAbsolute, ColumnAbsolute, ReferenceStyle, External)` --
    // FALSE, FALSE, xlA1, TRUE, which is `[Book1]Sheet1!A1:B2`: the same
    // external form `callerref` uses, so an object argument and a calling cell
    // read alike.
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

    bool DescribeRange(IDispatch* r, char* out, int cap)
    {
        char addr[512];
        if (!RangeAddress(r, addr, sizeof(addr))) return false;

        long cells = 0;
        const bool haveCount = GetLong(r, L"Count", cells);

        // THE ADDRESS ALONE IS A COMPLETE ANSWER. The contents are extra, and
        // are declined rather than truncated when there are too many -- or when
        // the count could not be had, because a count we do not know is not a
        // count we may assume is small.
        if (!haveCount || cells > kMaxCellsToRead)
        {
            _snprintf_s(out, cap, _TRUNCATE, "(%s)", addr);
            return true;
        }

        VARIANT v;
        if (!GetProp(r, L"Value2", v, nullptr, 0))
        {
            _snprintf_s(out, cap, _TRUNCATE, "(%s)", addr);
            return true;
        }
        if (t_hold) t_hold->var = &v;
        // The same decoder the columns use, reading the VARIANT we hold. Sized to fit after
        // "(addr)=", so the value's own truncation marker survives.
        char val[4096]; val[0] = 0;
        const int room = cap - static_cast<int>(strlen(addr)) - 4;
        const int valCap = room < static_cast<int>(sizeof(val)) ? room : static_cast<int>(sizeof(val));
        const char* held = "";
        const bool got = valCap >= 16 &&
            DescribeVariantValue(reinterpret_cast<std::uint64_t>(&v), val, valCap, &held);
        if (t_hold) t_hold->var = nullptr;
        VariantClear(&v);

        if (got && val[0]) _snprintf_s(out, cap, _TRUNCATE, "(%s)=%s", addr, val);
        else               _snprintf_s(out, cap, _TRUNCATE, "(%s)", addr);
        return true;
    }

    bool DescribeWorksheet(IDispatch* ws, char* out, int cap)
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
        // `[Book1]Sheet1`, the shape callerref uses. Without the book it is the
        // sheet alone -- true, and less than we wanted, rather than a book name
        // invented to fill the brackets.
        if (book[0]) _snprintf_s(out, cap, _TRUNCATE, "([%s]%s)", book, sheet);
        else         _snprintf_s(out, cap, _TRUNCATE, "(%s)", sheet);
        return true;
    }

    bool DescribeWorkbook(IDispatch* wb, char* out, int cap)
    {
        char book[256];
        if (!GetStr(wb, L"Name", book, sizeof(book))) return false;
        _snprintf_s(out, cap, _TRUNCATE, "([%s])", book);
        return true;
    }

    // The whole of the work, with no object that needs unwinding, so the caller
    // can wrap it in SEH.
    bool DescribeInner(IDispatch* d, char* cls, int clsCap, char* detail, int detailCap)
    {
        cls[0] = 0; detail[0] = 0;

        // A class identified by QueryInterface is named by that identification. The type info
        // would say `_Worksheet`, since Excel's sheet object yields no coclass.
        bool detailed = false;
        if (Answers(d, kIidRange))
        {
            strncpy_s(cls, clsCap, "Range", _TRUNCATE);
            detailed = DescribeRange(d, detail, detailCap);
        }
        else if (Answers(d, kIidWorksheet))
        {
            strncpy_s(cls, clsCap, "Worksheet", _TRUNCATE);
            detailed = DescribeWorksheet(d, detail, detailCap);
        }
        else if (Answers(d, kIidWorkbook))
        {
            strncpy_s(cls, clsCap, "Workbook", _TRUNCATE);
            detailed = DescribeWorkbook(d, detail, detailCap);
        }
        else if (!TypeName(d, cls, clsCap) || !cls[0])
        {
            // Not one we know, and it will not say what it is. That is the whole
            // of the answer, and the caller renders the address.
            InterlockedIncrement64(&g_unknown);
            return false;
        }

        if (detailed) InterlockedIncrement64(&g_described);
        else        { detail[0] = 0; InterlockedIncrement64(&g_namedOnly); }
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


bool DescribeObjectDetail(std::uint64_t ptr, char* cls, int clsCap,
                          char* detail, int detailCap)
{
    if (!ptr || clsCap < 8 || detailCap < 8) return false;
    cls[0] = 0; detail[0] = 0;
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
        const bool ok = DescribeInner(d, cls, clsCap, detail, detailCap);
        Unhold(d); d->Release();
        t_hold = nullptr;
        return ok;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hold.ReleaseAll();
        t_hold = nullptr;
        // A COM call that faults costs this description and nothing else: the
        // caller renders `object@0x...`, which is what it did before any of this.
        cls[0] = 0; detail[0] = 0;
        return false;
    }
}
}   // namespace vba
