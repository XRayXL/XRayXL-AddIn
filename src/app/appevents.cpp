#include "appevents.h"
#include "core/clock.h"
#include "core/excel_om.h"
#include "core/log.h"
#include "core/textbuf.h"
#include "core/tracemodes.h"
#include "core/valueformat.h"
#include "emit/csv.h"
#include "vba/vbaobject.h"
#include "vba/vbaretdecode.h"

#include <windows.h>
#include <oaidl.h>
#include <ocidl.h>
#include <cstdio>
#include <cstring>
#include <initializer_list>
#include <sstream>
#include <vector>

namespace app
{
namespace appevents
{
    namespace
    {
        // DIID_AppEvents, from EXCEL.EXE's type library.
        const GUID kDiidAppEvents = { 0x00024413, 0, 0, { 0xC0,0,0,0,0,0,0,0x46 } };

        struct Param { char name[48]; char type[64]; };
        struct EventType
        {
            long  dispid;
            char  name[64];
            int   count;
            Param params[8];
        };

        std::vector<EventType> g_types;     // this Excel's events, read at Start or Probe
        core::events::Mask     g_available = 0;
        bool                   g_availableKnown = false;

        IConnectionPoint* g_cp = nullptr;
        DWORD             g_cookie = 0;
        IDispatch*        g_sink = nullptr;
        bool              g_live = false;
        core::events::Mask g_selected = 0;
        bool              g_recordUnknown = false;   // only when every known event is chosen

        volatile LONG64 g_recorded = 0, g_unknown = 0, g_faults = 0;
        core::TextBuf   g_args;              // main thread only

        void Narrow(const wchar_t* w, char* out, int cap)
        {
            if (!w || WideCharToMultiByte(CP_UTF8, 0, w, -1, out, cap, nullptr, nullptr) <= 0) out[0] = 0;
        }

        void TypeName(ITypeInfo* ti, const TYPEDESC& td, char* out, int cap)
        {
            // ByRef is how Excel passes Cancel; the row names the type, not the passing.
            if (td.vt == VT_PTR && td.lptdesc) { TypeName(ti, *td.lptdesc, out, cap); return; }
            if (td.vt == VT_USERDEFINED)
            {
                ITypeInfo* r = nullptr;
                if (SUCCEEDED(ti->GetRefTypeInfo(td.hreftype, &r)) && r)
                {
                    BSTR n = nullptr;
                    if (SUCCEEDED(r->GetDocumentation(MEMBERID_NIL, &n, nullptr, nullptr, nullptr)) && n)
                    {
                        Narrow(n, out, cap);
                        SysFreeString(n);
                    }
                    r->Release();
                    if (out[0]) return;
                }
            }
            const char* t;
            switch (td.vt)
            {
            case VT_BOOL:     t = "Boolean"; break;
            case VT_BSTR:     t = "String";  break;
            case VT_I2:       t = "Integer"; break;
            case VT_I4:       t = "Long";    break;
            case VT_R8:       t = "Double";  break;
            case VT_DISPATCH: t = "Object";  break;
            case VT_UNKNOWN:  t = "Unknown"; break;
            case VT_VARIANT:  t = "Variant"; break;
            default:          t = "?";       break;
            }
            _snprintf_s(out, cap, _TRUNCATE, "%s", t);
        }

        // Developer instrument: pretend this Excel lacks the named events, so the missing-event path
        // can be tested on an Excel that has them all. XRAYXL_DIAG only.
        bool AbsentForTest(const char* name)
        {
            if (!core::modes::DiagEnabled()) return false;
            char list[512];
            if (!GetEnvironmentVariableA("XRAYXL_EVENTS_ABSENT", list, sizeof(list))) return false;
            char* ctx = nullptr;
            for (char* t = strtok_s(list, ",", &ctx); t; t = strtok_s(nullptr, ",", &ctx))
                if (_stricmp(t, name) == 0) return true;
            return false;
        }

        // Which events this Excel has, and their parameters, from the running Excel's own type
        // library, so an event added later is named without a code change.
        bool ReadTypes(IDispatch* app, std::ostringstream& log)
        {
            g_types.clear();
            g_available = 0;
            g_availableKnown = false;
            ITypeInfo* ti = nullptr;
            if (FAILED(app->GetTypeInfo(0, LOCALE_USER_DEFAULT, &ti)) || !ti) { log << "no type info"; return false; }
            ITypeLib* lib = nullptr; UINT idx = 0;
            const HRESULT hrLib = ti->GetContainingTypeLib(&lib, &idx);
            ti->Release();
            if (FAILED(hrLib) || !lib) { log << "no type library"; return false; }
            ITypeInfo* ev = nullptr;
            const HRESULT hrEv = lib->GetTypeInfoOfGuid(kDiidAppEvents, &ev);
            lib->Release();
            if (FAILED(hrEv) || !ev) { log << "no AppEvents in the type library"; return false; }

            TYPEATTR* ta = nullptr;
            if (FAILED(ev->GetTypeAttr(&ta)) || !ta) { ev->Release(); log << "no AppEvents attributes"; return false; }
            for (UINT f = 0; f < ta->cFuncs; ++f)
            {
                FUNCDESC* fd = nullptr;
                if (FAILED(ev->GetFuncDesc(f, &fd)) || !fd) continue;
                // IUnknown and IDispatch's own methods are listed too, restricted.
                if (!(fd->wFuncFlags & FUNCFLAG_FRESTRICTED))
                {
                    EventType e = {};
                    e.dispid = fd->memid;
                    BSTR names[9] = {};
                    UINT got = 0;
                    const UINT want = static_cast<UINT>(fd->cParams + 1 > 9 ? 9 : fd->cParams + 1);
                    ev->GetNames(fd->memid, names, want, &got);
                    Narrow(got ? names[0] : nullptr, e.name, sizeof(e.name));
                    e.count = fd->cParams > 8 ? 8 : fd->cParams;
                    for (int p = 0; p < e.count; ++p)
                    {
                        if (static_cast<UINT>(p + 1) < got) Narrow(names[p + 1], e.params[p].name, sizeof(e.params[p].name));
                        if (!e.params[p].name[0]) _snprintf_s(e.params[p].name, _TRUNCATE, "p%d", p + 1);
                        TypeName(ev, fd->lprgelemdescParam[p].tdesc, e.params[p].type, sizeof(e.params[p].type));
                    }
                    for (UINT n = 0; n < got; ++n) SysFreeString(names[n]);
                    if (e.name[0] && !AbsentForTest(e.name))
                    {
                        g_types.push_back(e);
                        const int k = core::events::FindDispid(e.dispid);
                        if (k >= 0) g_available |= core::events::Mask(1) << k;
                    }
                }
                ev->ReleaseFuncDesc(fd);
            }
            ev->ReleaseTypeAttr(ta);
            ev->Release();
            g_availableKnown = true;
            return true;
        }

        const EventType* TypeOf(long dispid)
        {
            for (const EventType& e : g_types) if (e.dispid == dispid) return &e;
            return nullptr;
        }

        // Objects are always described: events run on the main thread, where the object model is
        // safe to call, and a row that cannot say which sheet changed says little.
        void WriteValue(VARIANT* v, core::ValueWriter& w)
        {
            if (v->vt == (VT_BYREF | VT_VARIANT) && v->pvarVal) v = v->pvarVal;
            const VARTYPE vt = static_cast<VARTYPE>(v->vt & ~VT_BYREF);
            const bool byRef = (v->vt & VT_BYREF) != 0;
            if (vt == VT_DISPATCH || vt == VT_UNKNOWN)
            {
                IUnknown* u = byRef ? (v->ppunkVal ? *v->ppunkVal : nullptr) : v->punkVal;
                if (!u) { w.Word("Nothing"); return; }
                const std::uint64_t ptr = reinterpret_cast<std::uint64_t>(u);
                if (!vba::DescribeObjectDetail(ptr, w)) { w.BeginObject(nullptr, ptr, nullptr); w.EndObject(); }
                return;
            }
            if (!vba::DescribeVariantValue(reinterpret_cast<std::uint64_t>(v), w)) w.Marker("?");
        }

        // Parameter `i`'s value. Excel names every event argument by its position; unnamed ones
        // arrive last first, as IDispatch::Invoke defines.
        VARIANT* ArgAt(DISPPARAMS* dp, UINT i)
        {
            for (UINT n = 0; n < dp->cNamedArgs; ++n)
                if (dp->rgdispidNamedArgs[n] == static_cast<DISPID>(i)) return &dp->rgvarg[n];
            const UINT positional = dp->cArgs - dp->cNamedArgs;
            return (i < positional) ? &dp->rgvarg[dp->cArgs - 1 - i] : nullptr;
        }

        IUnknown* ObjectOf(VARIANT* v)
        {
            if (v->vt == (VT_BYREF | VT_VARIANT) && v->pvarVal) v = v->pvarVal;
            const VARTYPE vt = static_cast<VARTYPE>(v->vt & ~VT_BYREF);
            if (vt != VT_DISPATCH && vt != VT_UNKNOWN) return nullptr;
            return (v->vt & VT_BYREF) ? (v->ppunkVal ? *v->ppunkVal : nullptr) : v->punkVal;
        }

        void OnEvent(DISPID id, DISPPARAMS* dp)
        {
            if (!g_live) return;
            const int k = core::events::FindDispid(id);
            const bool want = (k >= 0) ? ((g_selected >> k) & 1) != 0 : g_recordUnknown;
            if (!want) return;
            // One the list left out is not this Excel's, even when it arrives: XRAYXL_EVENTS_ABSENT.
            if (k >= 0 && g_availableKnown && !((g_available >> k) & 1)) return;

            const long long t0 = core::QpcNow();
            const EventType* et = TypeOf(id);
            char unknownName[32];
            _snprintf_s(unknownName, _TRUNCATE, "Event0x%lX", static_cast<unsigned long>(id));
            const char* name = (k >= 0) ? core::events::At(k).name : (et ? et->name : unknownName);

            const UINT argc = dp ? dp->cArgs : 0;
            char callerref[600] = {};
            g_args.Clear();
            core::WithValueWriter(g_args, [&](core::ValueWriter& w)
            {
                w.BeginArgs();
                for (UINT i = 0; i < argc; ++i)
                {
                    VARIANT* v = ArgAt(dp, i);
                    if (!v) continue;
                    char pn[16]; _snprintf_s(pn, _TRUNCATE, "p%u", i + 1);
                    const bool typed = et && static_cast<int>(i) < et->count;
                    w.BeginNamedArg(typed ? et->params[i].name : pn, typed ? et->params[i].type : "?");
                    WriteValue(v, w);
                    w.EndArg();
                }
                w.EndArgs();
            });
            // The target, else the sheet, else the workbook: the most specific place the event names.
            for (const char* role : { "Target", "Sh", "Wb" })
            {
                if (callerref[0] || !et) break;
                for (int i = 0; i < et->count && i < static_cast<int>(argc); ++i)
                    if (strcmp(et->params[i].name, role) == 0)
                        if (VARIANT* v = ArgAt(dp, static_cast<UINT>(i)))
                            if (IUnknown* u = ObjectOf(v))
                                vba::ObjectWhere(reinterpret_cast<std::uint64_t>(u), callerref, sizeof(callerref));
            }
            WriteRow("Excel", name, callerref, g_args.Text(), t0);
            InterlockedIncrement64(k >= 0 ? &g_recorded : &g_unknown);
            // Our time, which lands inside any traced call that raised the event.
            core::TracerTicks() += core::QpcNow() - t0;
        }

        void OnEventGuarded(DISPID id, DISPPARAMS* dp)
        {
            __try { OnEvent(id, dp); }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                InterlockedIncrement64(&g_faults);
                emit::csv::ReleaseHeldByThisThread();
            }
        }

        // The subscriber Excel calls. It reads the arguments and never writes one, so a Cancel
        // stays whatever Excel or another handler made it.
        class Sink final : public IDispatch
        {
            volatile LONG m_ref = 1;
        public:
            HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
            {
                if (!ppv) return E_POINTER;
                if (riid == IID_IUnknown || riid == IID_IDispatch || riid == kDiidAppEvents)
                {
                    *ppv = static_cast<IDispatch*>(this);
                    AddRef();
                    return S_OK;
                }
                *ppv = nullptr;
                return E_NOINTERFACE;
            }
            ULONG STDMETHODCALLTYPE AddRef() override { return static_cast<ULONG>(InterlockedIncrement(&m_ref)); }
            ULONG STDMETHODCALLTYPE Release() override
            {
                const LONG n = InterlockedDecrement(&m_ref);
                if (n == 0) delete this;
                return static_cast<ULONG>(n);
            }
            HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* n) override { if (n) *n = 0; return S_OK; }
            HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo**) override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) override { return E_NOTIMPL; }
            HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* dp,
                                             VARIANT*, EXCEPINFO*, UINT*) override
            {
                OnEventGuarded(id, dp);
                return S_OK;
            }
        };
    }

    bool Start(core::events::Mask selected, std::string& why)
    {
        Stop();
        InterlockedExchange64(&g_recorded, 0);
        InterlockedExchange64(&g_unknown, 0);
        InterlockedExchange64(&g_faults, 0);

        std::ostringstream log;
        IDispatch* app = core::excelom::AcquireApplication(log);
        if (!app) { why = "no Excel Application object (" + log.str() + ")"; return false; }

        std::ostringstream tl;
        if (!ReadTypes(app, tl)) core::Log::Warning("events: this Excel's event list could not be read (" + tl.str() + ")");

        IConnectionPointContainer* cpc = nullptr;
        HRESULT hr = app->QueryInterface(IID_IConnectionPointContainer, reinterpret_cast<void**>(&cpc));
        app->Release();
        if (FAILED(hr) || !cpc) { why = "Application offers no events"; return false; }
        hr = cpc->FindConnectionPoint(kDiidAppEvents, &g_cp);
        cpc->Release();
        if (FAILED(hr) || !g_cp) { g_cp = nullptr; why = "Application has no AppEvents connection point"; return false; }

        g_sink = new Sink();
        hr = g_cp->Advise(g_sink, &g_cookie);
        if (FAILED(hr))
        {
            char h[64]; _snprintf_s(h, _TRUNCATE, "subscribing failed (0x%08lX)", static_cast<unsigned long>(hr));
            why = h;
            g_sink->Release(); g_sink = nullptr;
            g_cp->Release();   g_cp = nullptr;
            return false;
        }
        g_selected = selected;
        // All, as the dialog shows it: every event this Excel has, not every one the catalogue names.
        const core::events::Mask have = g_availableKnown ? g_available : core::events::AllMask();
        g_recordUnknown = (selected & have) == have;
        g_live = true;
        return true;
    }

    void Stop()
    {
        g_live = false;
        if (g_cp)
        {
            g_cp->Unadvise(g_cookie);
            g_cp->Release();
            g_cp = nullptr;
            g_cookie = 0;
        }
        if (g_sink) { g_sink->Release(); g_sink = nullptr; }
    }

    core::events::Mask Available(bool& known)
    {
        known = g_availableKnown;
        return g_availableKnown ? g_available : core::events::AllMask();
    }

    core::events::Mask Probe(bool& known)
    {
        // An Excel's events do not change while it runs, so once read is enough.
        if (g_availableKnown) return Available(known);
        std::ostringstream log;
        if (IDispatch* app = core::excelom::AcquireApplication(log))
        {
            ReadTypes(app, log);
            app->Release();
        }
        return Available(known);
    }

    bool ParamsText(int k, wchar_t* out, int cap)
    {
        if (cap < 1) return false;
        out[0] = 0;
        if (k < 0 || k >= core::events::Count()) return false;
        const EventType* et = TypeOf(core::events::At(k).dispid);
        if (!et) return false;
        int j = 0;
        for (int p = 0; p < et->count && j < cap - 1; ++p)
        {
            const int w = _snwprintf_s(out + j, cap - j, _TRUNCATE, L"%s%S As %S", p ? L", " : L"",
                                       et->params[p].name, et->params[p].type);
            if (w < 0) break;
            j += w;
        }
        if (!et->count) _snwprintf_s(out, cap, _TRUNCATE, L"no parameters");
        return true;
    }

    Totals ReadTotals()
    {
        Totals t;
        t.recorded = InterlockedCompareExchange64(&g_recorded, 0, 0);
        t.unknown  = InterlockedCompareExchange64(&g_unknown,  0, 0);
        t.faults   = InterlockedCompareExchange64(&g_faults,   0, 0);
        return t;
    }

    void WriteRow(const char* source, const char* function, const char* callerref,
                  const char* args, long long qpc, bool keep)
    {
        char tid[16], q[24];
        _snprintf_s(tid, _TRUNCATE, "%lu", GetCurrentThreadId());
        _snprintf_s(q, _TRUNCATE, "%lld", qpc);
        emit::csv::Row row;
        row.kind = "event";  row.source = source;
        row.span = "0";  row.parent = "0";  row.depth = "0";
        row.thread = tid;  row.qpc = q;
        row.function = function;  row.callerref = callerref ? callerref : "";
        row.args = args ? args : "";
        if (keep) emit::csv::WriteRowKept(row);
        else      emit::csv::WriteRow(row);
    }
}
}
