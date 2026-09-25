#include "vbaobject.h"
#include "vbaretdecode.h"
#include "core/sheetquote.h"
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
    // VBA's `_Collection` and the Scripting Runtime's `IDictionary`.
    const GUID kIidCollection = { 0xA4C46780, 0x499F, 0x101B, { 0xBB,0x78,0x00,0xAA,0x00,0x38,0x3C,0xBB } };
    const GUID kIidDictionary = { 0x42C642C1, 0x97E1, 0x11CF, { 0x97,0x8F,0x00,0xA0,0x24,0x63,0xE0,0x6F } };

    // So a class that stops being recognised is visible rather than silently absent.
    volatile LONG64 g_described = 0, g_namedOnly = 0, g_unknown = 0;

    // Interfaces and VARIANTs held inside the guarded call, so a fault part-way through can
    // still release them. The VARIANTs live here, in the guarded frame, because the frames that
    // filled them are gone by the time the handler runs. Trivially destructible, for SEH.
    struct ComHold
    {
        IUnknown* p[8];
        int       n;
        VARIANT   var[4];
        bool      used[4];
        void Init() { n = 0; for (bool& u : used) u = false; }
        void Add(IUnknown* u)  { if (u && n < 8) p[n++] = u; }
        void Drop(IUnknown* u) { for (int i = n - 1; i >= 0; --i) if (p[i] == u) { p[i] = p[--n]; return; } }
        VARIANT* NewVar()
        {
            for (int i = 0; i < 4; ++i)
                if (!used[i]) { used[i] = true; VariantInit(&var[i]); return &var[i]; }
            return nullptr;
        }
        void FreeVar(VARIANT* v)
        {
            for (int i = 0; i < 4; ++i)
                if (used[i] && v == &var[i]) { VariantClear(v); used[i] = false; return; }
        }
        void ReleaseAll()
        {
            while (n > 0) { IUnknown* u = p[--n]; if (u) u->Release(); }
            for (int i = 0; i < 4; ++i) if (used[i]) { VariantClear(&var[i]); used[i] = false; }
        }
    };
    __declspec(thread) ComHold* t_hold = nullptr;
    void Hold(IUnknown* u)   { if (t_hold) t_hold->Add(u); }
    void Unhold(IUnknown* u) { if (t_hold) t_hold->Drop(u); }
    // Null when every slot is taken; the caller then reads nothing.
    VARIANT* NewVar()        { return t_hold ? t_hold->NewVar() : nullptr; }
    void FreeVar(VARIANT* v) { if (t_hold && v) t_hold->FreeVar(v); }

    // Value2 of `A:A` would materialise a 25 MB VARIANT array inside a calculation.
    constexpr long kMaxCellsToRead = 4096;
    constexpr long kMaxItemsToRead = 4096;

    // The containers being walked on this thread: a Collection can hold itself, and every
    // level re-enters here on the hot-path stack.
    constexpr int kMaxWalkNest = 8;
    __declspec(thread) IUnknown* t_walking[kMaxWalkNest];
    __declspec(thread) int       t_walkDepth = 0;

    // False when `d` is already being walked, or the nesting is too deep to walk it.
    bool EnterWalk(IDispatch* d)
    {
        if (t_walkDepth >= kMaxWalkNest) return false;
        IUnknown* id = nullptr;
        if (FAILED(d->QueryInterface(IID_IUnknown, reinterpret_cast<void**>(&id))) || !id) return false;
        id->Release();   // only compared; `d` keeps the object alive
        for (int i = 0; i < t_walkDepth; ++i) if (t_walking[i] == id) return false;
        t_walking[t_walkDepth++] = id;
        return true;
    }
    void LeaveWalk() { if (t_walkDepth > 0) --t_walkDepth; }

    bool Answers(IDispatch* d, const GUID& iid)
    {
        if (!d) return false;
        void* p = nullptr;
        if (FAILED(d->QueryInterface(iid, &p)) || !p) return false;
        reinterpret_cast<IUnknown*>(p)->Release();
        return true;
    }

    // `d` as the dual interface `iid`, held; null when it does not answer.
    IDispatch* As(IDispatch* d, const GUID& iid)
    {
        void* p = nullptr;
        if (FAILED(d->QueryInterface(iid, &p)) || !p) return nullptr;
        IDispatch* r = reinterpret_cast<IDispatch*>(p);
        Hold(r);
        return r;
    }
    void Drop(IDispatch* d) { if (d) { Unhold(d); d->Release(); } }

    // `args` in declaration order; DISPPARAMS wants them reversed.
    bool InvokeGet(IDispatch* d, const wchar_t* name, WORD flags, VARIANT& out,
                   VARIANT* args, int argc)
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
                                   flags, &dp, &out, nullptr, nullptr));
    }

    bool GetProp(IDispatch* d, const wchar_t* name, VARIANT& out, VARIANT* args, int argc)
    {
        return InvokeGet(d, name, DISPATCH_PROPERTYGET, out, args, argc);
    }

    // Collection's Count and Dictionary's Keys are methods, not properties.
    bool CallGet(IDispatch* d, const wchar_t* name, VARIANT& out)
    {
        return InvokeGet(d, name, DISPATCH_METHOD | DISPATCH_PROPERTYGET, out, nullptr, 0);
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

    // Value2 of `A1,C1:D2` is A1's alone, so a range of several areas has no single value.
    bool OneArea(IDispatch* r)
    {
        IDispatch* areas = GetObj(r, L"Areas");
        if (!areas) return false;
        Hold(areas);
        long n = 0;
        const bool one = GetLong(areas, L"Count", n) && n == 1;
        Unhold(areas); areas->Release();
        return one;
    }

    // Contents are declined, not cut, when too many, when the count is unknown (an unknown count
    // may not be assumed small), or when there is more than one area.
    bool DescribeRange(IDispatch* r, std::uint64_t ptr, core::ValueWriter& w)
    {
        char addr[512];
        if (!RangeAddress(r, addr, sizeof(addr))) return false;
        w.BeginObject("Range", ptr, addr);

        long cells = 0;
        const bool haveCount = GetLong(r, L"Count", cells);
        if (haveCount && cells <= kMaxCellsToRead && OneArea(r))
        {
            VARIANT* v = NewVar();
            if (v && GetProp(r, L"Value2", *v, nullptr, 0))
                DescribeVariantValue(reinterpret_cast<std::uint64_t>(v), w);
            FreeVar(v);
        }
        w.EndObject();
        return true;
    }

    bool CountOf(IDispatch* d, long& n)
    {
        VARIANT v;
        if (!CallGet(d, L"Count", v)) return false;
        bool ok = false;
        if (v.vt == VT_I4) { n = v.lVal; ok = n >= 0; }
        VariantClear(&v);
        return ok;
    }

    // One element as a Variant array's element reads; `?` for one that will not decode.
    void Element(VARIANT* v, core::ValueWriter& w)
    {
        const core::ValueWriter::Mark m = w.Save();
        if (!DescribeVariantValue(reinterpret_cast<std::uint64_t>(v), w))
        {
            w.Restore(m);
            w.Marker("?");
        }
    }

    // The items as a Variant array from 1, walked the way For Each walks them. False, having
    // written nothing, when the enumerator will not give exactly `n`.
    bool CollectionItems(IDispatch* c, long n, core::ValueWriter& w)
    {
        VARIANT* e = NewVar();
        if (!e) return false;
        DISPPARAMS none = { nullptr, nullptr, 0, 0 };
        IEnumVARIANT* it = nullptr;
        if (SUCCEEDED(c->Invoke(DISPID_NEWENUM, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD | DISPATCH_PROPERTYGET, &none, e, nullptr, nullptr)) &&
            (e->vt == VT_UNKNOWN || e->vt == VT_DISPATCH) && e->punkVal)
            e->punkVal->QueryInterface(IID_IEnumVARIANT, reinterpret_cast<void**>(&it));
        FreeVar(e);
        if (!it) return false;
        Hold(it);
        VARIANT* v = NewVar();
        if (!v) { Unhold(it); it->Release(); return false; }

        const core::ValueWriter::Mark m = w.Save();
        const long long lo[1] = { 1 }, hi[1] = { n };
        w.BeginArray("Variant", 1, lo, hi);
        w.BeginLevel();
        long got = 0;
        for (; got < n && !w.Full(); ++got)
        {
            ULONG f = 0;
            if (it->Next(1, v, &f) != S_OK || f != 1) break;
            Element(v, w);
            VariantClear(v);
        }
        FreeVar(v);
        Unhold(it); it->Release();
        w.EndLevel();
        w.EndArray();
        if (got < n && !w.Full()) { w.Restore(m); return false; }
        return true;
    }

    bool DescribeCollection(IDispatch* d, std::uint64_t ptr, core::ValueWriter& w)
    {
        IDispatch* c = As(d, kIidCollection);
        if (!c) return false;
        long n = 0;
        const bool counted = CountOf(c, n);
        if (counted)
        {
            w.BeginObject("Collection", ptr, nullptr);
            // Too many, or already being walked: named, with no contents.
            if (n <= kMaxItemsToRead && EnterWalk(c))
            {
                CollectionItems(c, n, w);
                LeaveWalk();
            }
            w.EndObject();
        }
        Drop(c);
        return counted;
    }

    // A 1-D Variant array of exactly `n`, as Keys and Items return.
    SAFEARRAY* VariantVector(const VARIANT& v, long n, LONG& lb)
    {
        if (v.vt != (VT_ARRAY | VT_VARIANT) || !v.parray) return nullptr;
        SAFEARRAY* sa = v.parray;
        LONG ub = 0;
        if (SafeArrayGetDim(sa) != 1 || FAILED(SafeArrayGetLBound(sa, 1, &lb)) ||
            FAILED(SafeArrayGetUBound(sa, 1, &ub)) || static_cast<long long>(ub) - lb + 1 != n)
            return nullptr;
        return sa;
    }

    // One row per entry, `{key,item}`, in Keys' order and from its lower bound. False, having
    // written nothing, when Keys and Items do not both give `n`.
    bool DictionaryEntries(IDispatch* dict, long n, core::ValueWriter& w)
    {
        VARIANT* keys  = NewVar();
        VARIANT* items = NewVar();
        const bool got = keys && items && CallGet(dict, L"Keys", *keys) && CallGet(dict, L"Items", *items);

        bool ok = false;
        LONG klb = 0, ilb = 0;
        SAFEARRAY* ks = got ? VariantVector(*keys, n, klb) : nullptr;
        SAFEARRAY* is = ks ? VariantVector(*items, n, ilb) : nullptr;
        if (is)
        {
            const long long lo[2] = { klb, 0 }, hi[2] = { static_cast<long long>(klb) + n - 1, 1 };
            const std::uint64_t extent[2] = { static_cast<std::uint64_t>(n), 2 };
            w.BeginArray("Variant", 2, lo, hi);
            core::WalkRowMajor(w, 2, extent, [&](const std::uint64_t* idx)
            {
                SAFEARRAY* sa = idx[1] == 0 ? ks : is;
                LONG i = static_cast<LONG>(idx[0]) + (idx[1] == 0 ? klb : ilb);
                VARIANT* p = nullptr;
                if (SUCCEEDED(SafeArrayPtrOfIndex(sa, &i, reinterpret_cast<void**>(&p))) && p)
                    Element(p, w);
                else
                    w.Marker("?");
            });
            w.EndArray();
            ok = true;
        }
        FreeVar(items);
        FreeVar(keys);
        return ok;
    }

    bool DescribeDictionary(IDispatch* d, std::uint64_t ptr, core::ValueWriter& w)
    {
        IDispatch* dict = As(d, kIidDictionary);
        if (!dict) return false;
        long n = 0;
        const bool counted = CountOf(dict, n);
        if (counted)
        {
            w.BeginObject("Dictionary", ptr, nullptr);
            if (n <= kMaxItemsToRead && EnterWalk(dict))
            {
                DictionaryEntries(dict, n, w);
                LeaveWalk();
            }
            w.EndObject();
        }
        Drop(dict);
        return counted;
    }

    // `[Book1]Sheet1`, quoted as a cell's callerref is; without the book, the sheet alone rather
    // than a guess.
    bool WorksheetWhere(IDispatch* ws, char* where, int cap)
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
        char plain[600];
        if (book[0]) _snprintf_s(plain, _TRUNCATE, "[%s]%s", book, sheet);
        else         _snprintf_s(plain, _TRUNCATE, "%s", sheet);
        core::QuoteSheetPrefix(plain, where, cap);
        return true;
    }

    bool WorkbookWhere(IDispatch* wb, char* where, int cap)
    {
        char book[256];
        if (!GetStr(wb, L"Name", book, sizeof(book))) return false;
        _snprintf_s(where, cap, _TRUNCATE, "[%s]", book);
        return true;
    }

    bool DescribeWorksheet(IDispatch* ws, std::uint64_t ptr, core::ValueWriter& w)
    {
        char where[600];
        if (!WorksheetWhere(ws, where, sizeof(where))) return false;
        w.BeginObject("Worksheet", ptr, where);
        w.EndObject();
        return true;
    }

    bool DescribeWorkbook(IDispatch* wb, std::uint64_t ptr, core::ValueWriter& w)
    {
        char where[300];
        if (!WorkbookWhere(wb, where, sizeof(where))) return false;
        w.BeginObject("Workbook", ptr, where);
        w.EndObject();
        return true;
    }

    bool WhereInner(IDispatch* d, char* out, int cap)
    {
        if (Answers(d, kIidRange))     return RangeAddress(d, out, cap);
        if (Answers(d, kIidWorksheet)) return WorksheetWhere(d, out, cap);
        if (Answers(d, kIidWorkbook))  return WorkbookWhere(d, out, cap);
        return false;
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
        else if (Answers(d, kIidCollection))
        {
            known = "Collection";
            detailed = DescribeCollection(d, ptr, w);
        }
        else if (Answers(d, kIidDictionary))
        {
            known = "Dictionary";
            detailed = DescribeDictionary(d, ptr, w);
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
    // A container's element is described from inside its container's call, so this nests.
    ComHold* const outer = t_hold;
    const int walkDepth = t_walkDepth;
    ComHold hold;
    hold.Init();
    t_hold = &hold;
    __try
    {
        // Callers pass VT_UNKNOWN too; QueryInterface is the only safe call on a bare IUnknown.
        IDispatch* d = nullptr;
        IUnknown*  unk = reinterpret_cast<IUnknown*>(ptr);
        if (FAILED(unk->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&d))) || !d)
        {
            t_hold = outer;
            return false;
        }
        Hold(d);
        const bool ok = DescribeInner(d, ptr, w);
        Unhold(d); d->Release();
        t_hold = outer;
        return ok;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hold.ReleaseAll();
        t_hold = outer;
        t_walkDepth = walkDepth;
        w.Restore(mark);
        return false;
    }
}

bool ObjectWhere(std::uint64_t ptr, char* out, int cap)
{
    if (!ptr || cap < 1) return false;
    out[0] = 0;
    ComHold* const outer = t_hold;
    ComHold hold;
    hold.Init();
    t_hold = &hold;
    __try
    {
        IDispatch* d = nullptr;
        IUnknown*  unk = reinterpret_cast<IUnknown*>(ptr);
        if (FAILED(unk->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&d))) || !d)
        {
            t_hold = outer;
            return false;
        }
        Hold(d);
        const bool ok = WhereInner(d, out, cap);
        Unhold(d); d->Release();
        t_hold = outer;
        if (!ok) out[0] = 0;
        return ok;
    }
    __except (EXCEPTION_EXECUTE_HANDLER)
    {
        hold.ReleaseAll();
        t_hold = outer;
        out[0] = 0;
        return false;
    }
}
}   // namespace vba
