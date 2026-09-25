#include "ribbon.h"
#include "ribbonmodel.h"
#include "ribbonart.h"
#include "optionsdlg.h"
#include "diagnosticsdlg.h"
#include "traceactions.h"

#include "app/session.h"
#include "core/contained.h"
#include "core/crashlog.h"
#include "core/excel_api.h"
#include "core/excel_om.h"
#include "core/log.h"
#include "core/notify.h"
#include "core/tracemodes.h"
#include "emit/csv.h"

#include <windows.h>
#include <oaidl.h>
#include <atomic>
#include <cstdio>
#include <cwchar>
#include <new>
#include <sstream>
#include <string>

using core::modes::Depth;
using core::modes::Param;
using core::modes::Source;

namespace
{
// ---- identity -------------------------------------------------------------

// {730E8300-14A0-45D3-AE1D-3EED7DDC3B30}
const CLSID kClsid =
    { 0x730E8300, 0x14A0, 0x45D3, { 0xAE, 0x1D, 0x3E, 0xED, 0x7D, 0xDC, 0x3B, 0x30 } };
const wchar_t* const kClsidText = L"{730E8300-14A0-45D3-AE1D-3EED7DDC3B30}";
const wchar_t* const kProgId    = L"XRayXL.RibbonUI";

// The two Office interfaces, declared here so no Office headers are needed.
struct __declspec(uuid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744"))
IDTExtensibility2 : public IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE OnConnection(IDispatch* App, int ConnectMode, IDispatch* AddInInst, SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnDisconnection(int RemoveMode, SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY** custom) = 0;
    virtual HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY** custom) = 0;
};

struct __declspec(uuid("000C0396-0000-0000-C000-000000000046"))
IRibbonExtensibility : public IDispatch
{
    virtual HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR RibbonID, BSTR* pbstrRibbonXML) = 0;
};

// ---- faults: an exception escaping a COM method makes combase kill the process ----
int NoteComFault(EXCEPTION_POINTERS* xp, const char* what)
{
    return core::contained::Note(xp, "RIBBON COM METHOD", what,
                                 "contained; COM would have killed the process here");
}

// The handler half: releases what the faulting thread held, then writes what the filter formatted.
void ReportComFault()
{
    app::ReleaseHeldByThisThread();
    core::contained::Report();
}

// ---- small IDispatch helpers (local, so this file stays deletable in one piece) ----

IDispatch* GetDispProp(IDispatch* obj, const wchar_t* name)
{
    if (!obj) return nullptr;
    DISPID id = 0;
    OLECHAR* n = const_cast<OLECHAR*>(name);
    if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return nullptr;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT out; VariantInit(&out);
    const HRESULT hr = obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &dp, &out, nullptr, nullptr);
    IDispatch* r = nullptr;
    if (SUCCEEDED(hr) && out.vt == VT_DISPATCH && out.pdispVal) { r = out.pdispVal; r->AddRef(); }
    VariantClear(&out);
    return r;
}

bool GetIntProp(IDispatch* obj, const wchar_t* name, int& out)
{
    if (!obj) return false;
    DISPID id = 0;
    OLECHAR* n = const_cast<OLECHAR*>(name);
    if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return false;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT v; VariantInit(&v);
    const HRESULT hr = obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &dp, &v, nullptr, nullptr);
    bool ok = false;
    if (SUCCEEDED(hr))
    {
        VARIANT i4; VariantInit(&i4);
        if (SUCCEEDED(VariantChangeType(&i4, &v, 0, VT_I4))) { out = i4.lVal; ok = true; }
        VariantClear(&i4);
    }
    VariantClear(&v);
    return ok;
}

bool PutVariantProp(IDispatch* obj, const wchar_t* name, VARIANT& v)
{
    if (!obj) return false;
    DISPID id = 0;
    OLECHAR* n = const_cast<OLECHAR*>(name);
    if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return false;
    DISPID put = DISPID_PROPERTYPUT;
    DISPPARAMS dp{ &v, &put, 1, 1 };
    return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYPUT, &dp, nullptr, nullptr, nullptr));
}

bool PutBoolProp(IDispatch* obj, const wchar_t* name, bool value)
{
    VARIANT v; VariantInit(&v);
    v.vt = VT_BOOL; v.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return PutVariantProp(obj, name, v);
}

bool PutIntProp(IDispatch* obj, const wchar_t* name, int value)
{
    VARIANT v; VariantInit(&v);
    v.vt = VT_I4; v.lVal = value;
    return PutVariantProp(obj, name, v);
}

// Arming goes through Application.Run, not in process: xlfGetDef needs the macro context a
// registered command has and a ribbon callback does not, or names fall back to the export.
bool RunCommand(const wchar_t* command)
{
    std::ostringstream om;
    IDispatch* app = core::excelom::AcquireApplication(om);
    if (!app) return false;
    DISPID id = 0;
    OLECHAR  nm[] = L"Run";
    OLECHAR* n = nm;
    bool ok = false;
    if (SUCCEEDED(app->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id)))
    {
        VARIANT arg; VariantInit(&arg);
        arg.vt = VT_BSTR; arg.bstrVal = SysAllocString(command);
        DISPPARAMS dp{ &arg, nullptr, 1, 0 };
        ok = SUCCEEDED(app->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_METHOD, &dp, nullptr, nullptr, nullptr));
        VariantClear(&arg);
    }
    app->Release();
    return ok;
}

bool InvokeNoArgs(IDispatch* obj, const wchar_t* name)
{
    if (!obj) return false;
    DISPID id = 0;
    OLECHAR* n = const_cast<OLECHAR*>(name);
    if (FAILED(obj->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return false;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_METHOD, &dp, nullptr, nullptr, nullptr));
}

// COMAddIns.Item("<progid>") -- AddRef'd, or nullptr.
IDispatch* CallItem(IDispatch* coll, const wchar_t* key)
{
    if (!coll) return nullptr;
    DISPID id = 0;
    OLECHAR  nm[] = L"Item";
    OLECHAR* n = nm;
    if (FAILED(coll->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return nullptr;
    VARIANT arg; VariantInit(&arg);
    arg.vt = VT_BSTR; arg.bstrVal = SysAllocString(key);
    DISPPARAMS dp{ &arg, nullptr, 1, 0 };
    VARIANT out; VariantInit(&out);
    const HRESULT hr = coll->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                    DISPATCH_METHOD | DISPATCH_PROPERTYGET, &dp, &out, nullptr, nullptr);
    IDispatch* r = nullptr;
    if (SUCCEEDED(hr) && out.vt == VT_DISPATCH && out.pdispVal) { r = out.pdispVal; r->AddRef(); }
    VariantClear(&out);
    VariantClear(&arg);
    return r;
}

// Callback arguments arrive by value or by reference, so every read goes through these.
const VARIANT* Solid(const VARIANT* v)
{
    if (v && v->vt == (VT_VARIANT | VT_BYREF) && v->pvarVal) return v->pvarVal;
    return v;
}
IDispatch* AsDisp(const VARIANT* raw)
{
    const VARIANT* v = Solid(raw);
    if (!v) return nullptr;
    if (v->vt == VT_DISPATCH) return v->pdispVal;
    if (v->vt == (VT_DISPATCH | VT_BYREF) && v->ppdispVal) return *v->ppdispVal;
    return nullptr;
}

// The IRibbonControl Office hands every callback; its Id says which control fired.
bool ControlId(IDispatch* ctl, wchar_t* out, size_t cap)
{
    out[0] = 0;
    if (!ctl) return false;
    DISPID id = 0;
    OLECHAR  nm[] = L"Id";
    OLECHAR* n = nm;
    if (FAILED(ctl->GetIDsOfNames(IID_NULL, &n, 1, LOCALE_USER_DEFAULT, &id))) return false;
    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
    VARIANT v; VariantInit(&v);
    const HRESULT hr = ctl->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_PROPERTYGET, &dp, &v, nullptr, nullptr);
    bool ok = false;
    if (SUCCEEDED(hr) && v.vt == VT_BSTR && v.bstrVal)
    { wcsncpy_s(out, cap, v.bstrVal, _TRUNCATE); ok = true; }
    VariantClear(&v);
    return ok;
}

// What each control is lives in ui/ribbonmodel, where it can be tested without Excel.
namespace M = ui::ribbon::model;

// ---- the add-in object ----------------------------------------------------

std::atomic<long> g_objects{ 0 };

// Every IRibbonUI handed over (onLoad fires more than once), and our own COMAddIn wrapper.
IDispatch* g_addInInst = nullptr;

constexpr int kMaxRibbons = 16;
IDispatch*        g_ribbonUi[kMaxRibbons] = {};
int               g_ribbonCount = 0;
DWORD             g_uiThread = 0;
std::atomic<bool> g_dirty{ false };
std::atomic<bool> g_stopping{ false };
std::atomic<bool> g_connected{ false };   // Excel holds the add-in: set on connect, cleared on teardown

// deduplicated by pointer
void RememberRibbonUi(IDispatch* ui)
{
    if (!ui) return;
    for (int i = 0; i < g_ribbonCount; ++i) if (g_ribbonUi[i] == ui) return;
    if (g_ribbonCount >= kMaxRibbons)
    {
        core::Log::Warning("ribbon: more than 16 ribbons in this process; the newest is not tracked");
        return;
    }
    ui->AddRef();
    g_ribbonUi[g_ribbonCount++] = ui;
}

void ReleaseRibbonUis()
{
    // Cleared before releasing, so a re-entrant teardown finds nothing to do.
    const int n = g_ribbonCount;
    g_ribbonCount = 0;
    for (int i = 0; i < n; ++i)
    {
        IDispatch* ui = g_ribbonUi[i];
        g_ribbonUi[i] = nullptr;
        if (ui) ui->Release();
    }
}

// Callbacks Office has made: logged beside each invalidate, as evidence one reached a ribbon.
long g_callbacks = 0;

void InvalidateNow()
{
    g_dirty.store(false);
    const long before = g_callbacks;
    int sent = 0;
    for (int i = 0; i < g_ribbonCount; ++i)
        if (g_ribbonUi[i] && InvokeNoArgs(g_ribbonUi[i], L"Invalidate")) ++sent;
    if (core::Log::GetLevel() == core::Log::Level::Debug)
    {
        char line[144];
        _snprintf_s(line, _TRUNCATE,
                    "ribbon: invalidate -> %d of %d ribbon(s); %ld callback(s) before this one",
                    sent, g_ribbonCount, before);
        core::Log::Debug(line);
    }
}

// State changed elsewhere. Off the UI thread it only marks the display stale.
void OnStateChanged()
{
    if (g_stopping.load() || g_ribbonCount == 0) return;
    if (GetCurrentThreadId() == g_uiThread) InvalidateNow();
    else                                    g_dirty.store(true);
}

// The dispids are model::Callback, so a name can only be wrong in one place.
enum : DISPID {
    DISPID_ONLOAD     = M::CbOnLoad,     DISPID_ONARM     = M::CbOnArm,
    DISPID_ONDISARM   = M::CbOnDisarm,   DISPID_GETENABLED = M::CbGetEnabled,
    DISPID_ONOPTIONS  = M::CbOnOptions,  DISPID_LOADIMAGE = M::CbLoadImage,
    DISPID_ONDIAGNOSTICS = M::CbOnDiagnostics, DISPID_ONTAIL = M::CbOnTail
};

// defined with the connect plumbing below
HWND MainWindow();
void MarkConnected();

int  g_lastRegistered = 0;
int  g_lastArmed = 0;
std::atomic<bool> g_explainEmptyArm{ false };

// a leaf, because a std::string may not be in scope in a __try frame (C2712)
void ArmNoUnwind()
{
    xll::ArmReport r = app::Arm();
    g_lastRegistered = r.registered;
    g_lastArmed      = r.armed;
    core::crashlog::Note(r.detail.c_str());
}

// An arm that hooks nothing looks like a broken button, so it is explained -- from a timer.
void ExplainEmptyArm()
{
    HWND owner = MainWindow();
    {
        char d[128];
        _snprintf_s(d, _TRUNCATE, "ribbon: explain box -- owner=%p visible=%d",
                    (void*)owner, owner ? (int)IsWindowVisible(owner) : -1);
        core::Log::Debug(d);
    }
    if (!owner || !IsWindowVisible(owner)) return;

    // The one cause that is the user's to fix, because they did it on this tab.
    const bool vbaOff = !core::modes::VbaEnabled();

    wchar_t text[1000];
    if (vbaOff)
        _snwprintf_s(text, _TRUNCATE,
            L"There is nothing for XRayXL to record in this Excel yet.\r\n\r\n"
            L"VBA tracing is switched off in XRayXL's Options, and no other add-in "
            L"is loaded for XRayXL to follow.\r\n\r\n"
            L"Switch VBA tracing back on, or open the workbook or add-in you want "
            L"to look at, then press Arm again.");
    else
        _snwprintf_s(text, _TRUNCATE,
            L"There is nothing for XRayXL to record in this Excel yet.\r\n\r\n"
            L"XRayXL follows the add-in functions and VBA macros that run inside "
            L"Excel. This Excel has not opened any yet.\r\n\r\n"
            L"Open the workbook or load the add-in you want to look at, then press "
            L"Arm again.\r\n\r\n"
            L"The demo add-ins and workbooks are set up to show this working.");

    MessageBoxW(owner, text, L"XRayXL", MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND);
}

VOID CALLBACK ExplainTimerProc(HWND, UINT, UINT_PTR id, DWORD)
{
    KillTimer(nullptr, id);
    core::Log::Debug("ribbon: explain timer fired");
    if (g_explainEmptyArm.exchange(false)) ExplainEmptyArm();
}

class RibbonAddin : public IDTExtensibility2, public IRibbonExtensibility
{
public:
    RibbonAddin() { g_objects.fetch_add(1); }

    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        *ppv = nullptr;
        if (riid == IID_IUnknown || riid == IID_IDispatch || riid == __uuidof(IDTExtensibility2))
            *ppv = static_cast<IDTExtensibility2*>(this);
        else if (riid == __uuidof(IRibbonExtensibility))
            *ppv = static_cast<IRibbonExtensibility*>(this);
        else
            return E_NOINTERFACE;
        AddRef();
        return S_OK;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return static_cast<ULONG>(++m_ref); }
    ULONG STDMETHODCALLTYPE Release() override
    {
        const long n = --m_ref;
        if (n == 0) { g_objects.fetch_sub(1); delete this; return 0; }
        return static_cast<ULONG>(n < 0 ? 0 : n);
    }

    HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT* p) override { if (p) *p = 0; return S_OK; }
    HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT, LCID, ITypeInfo** p) override { if (p) *p = nullptr; return E_NOTIMPL; }

    HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID, LPOLESTR* names, UINT cNames, LCID, DISPID* ids) override
    {
        __try { return NamesImpl(names, cNames, ids); }
        __except (NoteComFault(GetExceptionInformation(), "GetIDsOfNames")) { ReportComFault(); return E_FAIL; }
    }

    HRESULT STDMETHODCALLTYPE Invoke(DISPID id, REFIID, LCID, WORD, DISPPARAMS* dp,
                                     VARIANT* result, EXCEPINFO*, UINT*) override
    {
        __try { return InvokeImpl(id, dp, result); }
        __except (NoteComFault(GetExceptionInformation(), "Invoke"))
        {
            // S_OK: Office reports a failing callback as a broken add-in.
            ReportComFault();
            return S_OK;
        }
    }

    // ---- IDTExtensibility2: AddInInst is kept for teardown; Application never is ----
    HRESULT STDMETHODCALLTYPE OnConnection(IDispatch*, int mode, IDispatch* AddInInst, SAFEARRAY**) override
    {
        // OnConnection fires more than once: release what is held before taking the new pointer.
        if (AddInInst != g_addInInst)
        {
            if (g_addInInst) g_addInInst->Release();
            if (AddInInst)   AddInInst->AddRef();
            g_addInInst = AddInInst;
        }
        char line[96];
        _snprintf_s(line, _TRUNCATE, "ribbon: OnConnection (ConnectMode %d)%s", mode,
                    AddInInst ? "" : " -- no add-in object offered");
        core::Log::Note(line);
        // Not always our Connect = True: Enable Content on the Security Warning bar lands here too.
        MarkConnected();
        return S_OK;
    }
    // RemoveMode 0 is Excel exiting, after any Cancel; 1 is a disconnect asked for.
    HRESULT STDMETHODCALLTYPE OnDisconnection(int mode, SAFEARRAY**) override
    {
        __try { return TeardownImpl("ribbon: OnDisconnection", mode); }
        __except (NoteComFault(GetExceptionInformation(), "OnDisconnection")) { ReportComFault(); return S_OK; }
    }
    HRESULT STDMETHODCALLTYPE OnAddInsUpdate(SAFEARRAY**) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnStartupComplete(SAFEARRAY**) override { return S_OK; }
    HRESULT STDMETHODCALLTYPE OnBeginShutdown(SAFEARRAY**) override
    {
        __try { return TeardownImpl("ribbon: OnBeginShutdown", -1); }
        __except (NoteComFault(GetExceptionInformation(), "OnBeginShutdown")) { ReportComFault(); return S_OK; }
    }

    // ---- IRibbonExtensibility ----
    HRESULT STDMETHODCALLTYPE GetCustomUI(BSTR, BSTR* xml) override
    {
        __try
        {
            if (!xml) return E_POINTER;
            *xml = SysAllocString(M::kCustomUi);
            core::crashlog::Note("ribbon: GetCustomUI served");
            return *xml ? S_OK : E_OUTOFMEMORY;
        }
        __except (NoteComFault(GetExceptionInformation(), "GetCustomUI")) { ReportComFault(); return E_FAIL; }
    }

private:
    HRESULT NamesImpl(LPOLESTR* names, UINT cNames, DISPID* ids)
    {
        HRESULT hr = S_OK;
        for (UINT i = 0; i < cNames; ++i)
        {
            const M::Callback cb = M::CallbackForName(names[i]);
            if (cb == M::CbUnknown) { ids[i] = DISPID_UNKNOWN; hr = DISP_E_UNKNOWNNAME; }
            else                     ids[i] = static_cast<DISPID>(cb);
        }
        return hr;
    }

    // DISPPARAMS holds the arguments backwards: rgvarg[0] is the last parameter.
    static const VARIANT* Arg(DISPPARAMS* dp, UINT fromEnd)
    {
        if (!dp || dp->cArgs <= fromEnd) return nullptr;
        return &dp->rgvarg[fromEnd];
    }
    // Every ribbon callback takes the IRibbonControl first, so it is last here.
    static const VARIANT* ControlArg(DISPPARAMS* dp)
    {
        if (!dp || dp->cArgs == 0) return nullptr;
        return &dp->rgvarg[dp->cArgs - 1];
    }

    HRESULT InvokeImpl(DISPID id, DISPPARAMS* dp, VARIANT* result)
    {
        if (result) VariantInit(result);

        if (id == DISPID_ONLOAD)
        {
            // onLoad(ribbon as IRibbonUI): held until teardown
            IDispatch* ui = AsDisp(Arg(dp, 0));
            if (ui) { g_uiThread = GetCurrentThreadId(); RememberRibbonUi(ui); }

            char line[80];
            _snprintf_s(line, _TRUNCATE, "ribbon: controls are live (%d ribbon(s) tracked)", g_ribbonCount);
            core::Log::Note(ui ? line : "ribbon: onLoad without an IRibbonUI");
            return S_OK;
        }

        if (id == DISPID_LOADIMAGE)
        {
            // loadImage(imageId): asked once per image id and cached by Office.
            // From the window chain, not the add-in object: the ribbon owns no lifetime state.
            const VARIANT* which = Arg(dp, 0);
            const wchar_t* name = (which && which->vt == VT_BSTR && which->bstrVal) ? which->bstrVal : L"";
            std::ostringstream om;
            IDispatch* app = core::excelom::AcquireApplication(om);
            IDispatch* pic = nullptr;
            std::string what = "an unknown image's";
            if      (_wcsicmp(name, L"arm") == 0)    { pic = ui::ribbon::art::ArmPicture(app, MainWindow());    what = "Arm's"; }
            else if (_wcsicmp(name, L"disarm") == 0) { pic = ui::ribbon::art::DisarmPicture(app, MainWindow()); what = "Disarm's"; }
            else if (_wcsicmp(name, L"tail") == 0)   { pic = ui::ribbon::art::TailPicture(app, MainWindow());   what = "Tail's"; }
            if (app) app->Release();
            if (!pic) core::Log::Warning("ribbon: " + what + " picture could not be made; the button shows no icon");
            else      core::Log::Debug("ribbon: " + what + " picture made");
            if (result && pic) { result->vt = VT_DISPATCH; result->pdispVal = pic; }
            else if (pic) pic->Release();
            return S_OK;
        }

        // everything below is a callback about a control Office is displaying
        ++g_callbacks;

        wchar_t cid[64] = {};
        ControlId(AsDisp(ControlArg(dp)), cid, 64);

        switch (id)
        {
        case DISPID_ONARM:
        {
            core::Log::Note("ribbon: Arm pressed");
            if (!app::IsArmed() && !RunCommand(L"XRayXL_Arm"))
            {
                core::Log::Warning("ribbon: Application.Run(\"XRayXL_Arm\") failed; arming in place,"
                                   " which leaves registered names unresolved");
                ArmNoUnwind();
            }
            // arming can legitimately hook nothing, and then nothing on the ribbon changes
            const bool armed = app::IsArmed();
            if (!armed)
            {
                char line[160];
                _snprintf_s(line, _TRUNCATE,
                            "ribbon: Arm hooked nothing (%d registered, %d armed)"
                            " -- telling the user why",
                            g_lastRegistered, g_lastArmed);
                core::Log::Warning(line);
                g_explainEmptyArm.store(true);
                SetTimer(nullptr, 0, 1, ExplainTimerProc);   // after this callback returns
            }
            InvalidateNow();
            return S_OK;
        }

        case DISPID_ONOPTIONS:
            ui::options::Show(MainWindow());
            InvalidateNow();        // Arm and Disarm follow whatever it changed
            return S_OK;

        case DISPID_ONDIAGNOSTICS:
            core::Log::Note("ribbon: Diagnostics pressed");
            ui::diagnostics::Show(MainWindow());
            return S_OK;

        case DISPID_ONTAIL:
        {
            core::Log::Note("ribbon: Tail pressed");
            const ui::trace::Result r = ui::trace::TailInPowerShell(emit::csv::Path());
            if (r != ui::trace::Result::Ok)
                MessageBoxW(MainWindow(), ui::trace::Explain(r), L"XRayXL", MB_OK | MB_ICONINFORMATION | MB_SETFOREGROUND);
            return S_OK;
        }

        case DISPID_ONDISARM:
            core::Log::Note("ribbon: Disarm pressed");
            if (app::IsArmed() && !RunCommand(L"XRayXL_Disarm")) app::Disarm();
            InvalidateNow();
            return S_OK;

        case DISPID_GETENABLED:
            if (result)
            {
                const bool live = M::EnabledFor(cid, app::IsArmed(), !emit::csv::Path().empty());
                result->vt = VT_BOOL;
                result->boolVal = live ? VARIANT_TRUE : VARIANT_FALSE;
            }
            return S_OK;

        default:
            return DISP_E_MEMBERNOTFOUND;
        }
    }

    // Both shutdown callbacks land here, so it runs twice safely. It also disarms: Excel
    // calls them only once the exit can no longer be cancelled, unlike xlAutoClose.
    HRESULT TeardownImpl(const char* which, int mode)
    {
        core::crashlog::Note(which);
        char line[128];
        char mode_[32] = "";
        if (mode >= 0) _snprintf_s(mode_, _TRUNCATE, " (RemoveMode %d)", mode);
        _snprintf_s(line, _TRUNCATE, "%s%s%s", which, mode_,
                    g_stopping.load() ? "" : " -- tearing down");
        core::Log::Note(line);
        // Also a disconnect mid-session (RemoveMode 1): after it, xlAutoClose must disarm itself.
        g_connected.store(false);
        core::SubscribeStateChanged(nullptr);
        ReleaseRibbonUis();
        if (g_addInInst) { IDispatch* a = g_addInInst; g_addInInst = nullptr; a->Release(); }
        if (!g_stopping.load()) app::Disarm(true);
        return S_OK;
    }

    std::atomic<long> m_ref{ 1 };
};


class ClassFactory : public IClassFactory
{
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override
    {
        if (!ppv) return E_POINTER;
        if (riid == IID_IUnknown || riid == IID_IClassFactory)
        { *ppv = static_cast<IClassFactory*>(this); AddRef(); return S_OK; }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return 2; }     // a static singleton
    ULONG STDMETHODCALLTYPE Release() override { return 1; }

    HRESULT STDMETHODCALLTYPE CreateInstance(IUnknown* outer, REFIID riid, void** ppv) override
    {
        __try { return CreateImpl(outer, riid, ppv); }
        __except (NoteComFault(GetExceptionInformation(), "CreateInstance")) { ReportComFault(); return E_UNEXPECTED; }
    }
    HRESULT STDMETHODCALLTYPE LockServer(BOOL) override { return S_OK; }

private:
    static HRESULT CreateImpl(IUnknown* outer, REFIID riid, void** ppv)
    {
        if (outer) return CLASS_E_NOAGGREGATION;
        RibbonAddin* obj = new (std::nothrow) RibbonAddin();
        if (!obj) return E_OUTOFMEMORY;
        const HRESULT hr = obj->QueryInterface(riid, ppv);
        obj->Release();
        return hr;
    }
};

ClassFactory g_factory;

// ---- transient registration ----------------------------------------------

LSTATUS SetSz(const std::wstring& subkey, const wchar_t* name, const wchar_t* value)
{
    HKEY k = nullptr;
    LSTATUS s = RegCreateKeyExW(HKEY_CURRENT_USER, subkey.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &k, nullptr);
    if (s != ERROR_SUCCESS) return s;
    s = RegSetValueExW(k, name, 0, REG_SZ, reinterpret_cast<const BYTE*>(value),
                       static_cast<DWORD>((wcslen(value) + 1) * sizeof(wchar_t)));
    RegCloseKey(k);
    return s;
}
LSTATUS SetDword(const std::wstring& subkey, const wchar_t* name, DWORD value)
{
    HKEY k = nullptr;
    LSTATUS s = RegCreateKeyExW(HKEY_CURRENT_USER, subkey.c_str(), 0, nullptr, 0, KEY_WRITE, nullptr, &k, nullptr);
    if (s != ERROR_SUCCESS) return s;
    s = RegSetValueExW(k, name, 0, REG_DWORD, reinterpret_cast<const BYTE*>(&value), sizeof(value));
    RegCloseKey(k);
    return s;
}

std::wstring AddinsKey() { return std::wstring(L"Software\\Microsoft\\Office\\Excel\\Addins\\") + kProgId; }
std::wstring ClsidKey()  { return std::wstring(L"Software\\Classes\\CLSID\\") + kClsidText; }
std::wstring ProgIdKey() { return std::wstring(L"Software\\Classes\\") + kProgId; }

// HKCU only. An elevated Excel ignores the user hive, which surfaces as a refused connect.
bool RegisterServer()
{
    const std::wstring xll = core::GetOwnModulePath();
    if (xll.empty()) return false;
    const std::wstring clsid = ClsidKey();
    bool ok = true;
    ok &= SetSz(clsid, nullptr, L"XRayXL RibbonUI") == ERROR_SUCCESS;
    ok &= SetSz(clsid + L"\\InprocServer32", nullptr, xll.c_str()) == ERROR_SUCCESS;
    ok &= SetSz(clsid + L"\\InprocServer32", L"ThreadingModel", L"Both") == ERROR_SUCCESS;
    ok &= SetSz(clsid + L"\\ProgID", nullptr, kProgId) == ERROR_SUCCESS;
    ok &= SetSz(ProgIdKey() + L"\\CLSID", nullptr, kClsidText) == ERROR_SUCCESS;
    wchar_t friendly[48];
    _snwprintf_s(friendly, _TRUNCATE, L"XRayXL Ribbon v%hs", app::VersionText());
    ok &= SetSz(AddinsKey(), L"FriendlyName", friendly) == ERROR_SUCCESS;
    ok &= SetSz(AddinsKey(), L"Description", L"XRayXL, Excel calculation diagnostics") == ERROR_SUCCESS;
    // LoadBehavior 0: Connect = True loads it anyway, and 3 would autoload it into the next Excel.
    ok &= SetDword(AddinsKey(), L"LoadBehavior", 0) == ERROR_SUCCESS;
    return ok;
}

void UnregisterServer()
{
    RegDeleteTreeW(HKEY_CURRENT_USER, AddinsKey().c_str());
    RegDeleteTreeW(HKEY_CURRENT_USER, ClsidKey().c_str());
    RegDeleteTreeW(HKEY_CURRENT_USER, ProgIdKey().c_str());
}

// ---- connect --------------------------------------------------------------

UINT_PTR g_timer     = 0;
int      g_attempt   = 0;   // real connect attempts -- not the waiting below
int      g_waits     = 0;   // polls spent waiting for an object model to exist

void UnregisterServer();

// Once, however the connect arrived. Excel built the object from the keys before calling
// OnConnection, so they have done their job.
void MarkConnected()
{
    if (g_stopping.load() || g_connected.exchange(true)) return;
    UnregisterServer();
    core::SubscribeStateChanged(OnStateChanged);
    core::Log::Info("ribbon: XRayXL buttons loaded");
}

// A visible XLMAIN of ours: the first in z-order is often an invisible container.
HWND MainWindow()
{
    const DWORD mine = GetCurrentProcessId();
    HWND firstOfOurs = nullptr;
    HWND h = nullptr;
    while ((h = FindWindowExW(nullptr, h, L"XLMAIN", nullptr)) != nullptr)
    {
        DWORD owner = 0;
        GetWindowThreadProcessId(h, &owner);
        if (owner != mine) continue;
        if (IsWindowVisible(h)) return h;
        if (!firstOfOurs) firstOfOurs = h;
    }
    return firstOfOurs;
}

BOOL CALLBACK LookForExcel7Child(HWND hwnd, LPARAM found)
{
    wchar_t cls[32] = {};
    GetClassNameW(hwnd, cls, 31);
    if (wcscmp(cls, L"EXCEL7") == 0) { *reinterpret_cast<bool*>(found) = true; return FALSE; }
    return TRUE;
}

// No workbook, no object model: Application is reached through a workbook's EXCEL7 window.
bool ObjectModelReachable()
{
    // every XLMAIN of ours: the one carrying the EXCEL7 child is not necessarily the first
    const DWORD mine = GetCurrentProcessId();
    HWND h = nullptr;
    while ((h = FindWindowExW(nullptr, h, L"XLMAIN", nullptr)) != nullptr)
    {
        DWORD owner = 0;
        GetWindowThreadProcessId(h, &owner);
        if (owner != mine) continue;
        bool found = false;
        EnumChildWindows(h, LookForExcel7Child, reinterpret_cast<LPARAM>(&found));
        if (found) return true;
    }
    return false;
}

bool MessageBoxOff()
{
    wchar_t v[8]{};
    return GetEnvironmentVariableW(L"XRAYXL_NOMESSAGEBOX", v, 8) > 0 && v[0] == L'1';
}

// A message box only where a user is watching: in any other Excel it would block the process.
void FailSoft(const wchar_t* reason, const char* logged)
{
    core::Log::Warning(std::string("ribbon: ") + logged +
                       " -- no ribbon buttons; XRayXL_Arm, XRayXL_Disarm and the"
                       " trace parameters are unaffected");
    core::crashlog::Note("ribbon: NOT loaded; the XLL commands are unaffected");

    HWND owner = MainWindow();
    const char* silent =
        !owner || !IsWindowVisible(owner) ? "Excel has no visible window"
        : MessageBoxOff()                 ? "XRAYXL_NOMESSAGEBOX=1"
        :                                   nullptr;
    if (silent)
    {
        core::Log::Note(std::string("ribbon: ") + silent + ", so no message box");
        return;
    }
    wchar_t text[1024];
    _snwprintf_s(text, _TRUNCATE,
        L"XRayXL could not load its ribbon.\n\n%s\n\n"
        L"XRayXL's functions and macros work without the ribbon. Arm and disarm from VBA, "
        L"the Macro dialog, or any automation client:\n\n"
        L"    Application.Run \"XRayXL_Arm\"\n"
        L"    Application.Run \"XRayXL_Disarm\"\n\n"
        L"Capture settings:\n"
        L"    Application.Run \"XRayXL_SetTraceParam\", \"XLL\", \"DEPTH\", \"TOP\"\n\n"
        L"Details are in the log, under %%TEMP%%\\XRayXL\\Logs.", reason);
    MessageBoxW(owner, text, L"XRayXL", MB_OK | MB_ICONWARNING | MB_SETFOREGROUND);
}

void ConnectNow();

VOID CALLBACK ConnectTimerProc(HWND, UINT, UINT_PTR id, DWORD)
{
    KillTimer(nullptr, id);
    g_timer = 0;
    ConnectNow();
}

bool Reschedule(UINT ms)
{
    g_timer = SetTimer(nullptr, 0, ms, ConnectTimerProc);
    return g_timer != 0;
}

// Waiting for a workbook: brisk at first, then every 2 s for about ten minutes.
constexpr int  kMaxWaits = 310;
UINT WaitDelayMs(int waits)
{
    if (waits < 10) return 200;      // first 2 s   -- a warm Excel lands here
    if (waits < 30) return 500;      // next 10 s   -- a cold start lands here
    return 2000;                     // then patient
}

// The usual failure is timing, so retry. True if another attempt is coming.
bool RetryOrFail(const wchar_t* reason, const char* logged)
{
    if (g_attempt < 3)
    {
        g_timer = SetTimer(nullptr, 0, 750, ConnectTimerProc);
        if (g_timer)
        {
            core::Log::Debug(std::string("ribbon: ") + logged + "; retrying");
            return true;
        }
    }
    FailSoft(reason, logged);
    return false;
}

void ConnectNow()
{
    if (g_connected.load() || g_stopping.load()) return;

    // No workbook yet is not a failure: wait for one.
    if (!ObjectModelReachable())
    {
        if (++g_waits <= kMaxWaits && Reschedule(WaitDelayMs(g_waits)))
        {
            if (g_waits == 1)
                core::Log::Note("ribbon: no workbook yet, so no object model to connect through -- waiting");
            return;
        }
        // No message box: there was never a ribbon to add to. The log says what happened.
        core::Log::Warning("ribbon: gave up waiting for a workbook; no ribbon buttons this session"
                           " -- XRayXL_Arm, XRayXL_Disarm and the trace parameters are unaffected");
        return;
    }

    ++g_attempt;

    std::ostringstream om;
    IDispatch* app = core::excelom::AcquireApplication(om);
    if (!om.str().empty()) core::Log::Debug("ribbon: " + om.str());
    if (!app)
    {
        // transient: the XLL can load before Excel has built that window
        RetryOrFail(L"Excel's object model could not be reached.", "no Application object");
        return;
    }

    if (!RegisterServer())
    {
        app->Release();
        UnregisterServer();
        FailSoft(L"XRayXL could not write its temporary registration under "
                 L"HKEY_CURRENT_USER. That is usually a locked-down or elevated session.",
                 "RegisterServer failed");
        return;
    }

    // macro security can refuse the connect; lowered for it and put straight back
    int oldSecurity = 0;
    const bool hadSecurity = GetIntProp(app, L"AutomationSecurity", oldSecurity);
    if (hadSecurity && oldSecurity != 1) PutIntProp(app, L"AutomationSecurity", 1);

    bool connected = false;
    if (IDispatch* coll = GetDispProp(app, L"COMAddIns"))
    {
        InvokeNoArgs(coll, L"Update");                  // make Excel read the key just written
        if (IDispatch* entry = CallItem(coll, kProgId))
        {
            // One pass. A second fires OnDisconnection on the real object, which disarms.
            connected = PutBoolProp(entry, L"Connect", true);
            entry->Release();
        }
        coll->Release();
    }

    if (hadSecurity && oldSecurity != 1) PutIntProp(app, L"AutomationSecurity", oldSecurity);
    app->Release();

    if (!connected)
    {
        // Keys kept: Enable Content on the Security Warning bar connects through them.
        // Excel also refuses while it is still loading; the retry covers that.
        RetryOrFail(L"This is most likely because of Excel's security settings. If Excel shows "
                    L"a yellow Security Warning bar, clicking Enable Content may load it.",
                    "COMAddIns connect refused");
        return;
    }

    MarkConnected();      // normally already done by OnConnection, keys and all
}

bool RibbonDisabledByEnv()
{
    wchar_t v[8]{};
    return GetEnvironmentVariableW(L"XRAYXL_RIBBON", v, 8) > 0 && v[0] == L'0';
}
}   // namespace

namespace ui
{
namespace ribbon
{
    void Start()
    {
        static volatile LONG s_started = 0;
        if (InterlockedExchange(&s_started, 1) != 0) return;   // xlAutoOpen can arrive twice

        if (RibbonDisabledByEnv())
        {
            core::Log::Info("ribbon: XRAYXL_RIBBON=0 -- no ribbon buttons this session");
            return;
        }

        // Excel refuses the connect from inside xlAutoOpen, so it runs from a timer.
        g_timer = SetTimer(nullptr, 0, 10, ConnectTimerProc);
        core::Log::Note(g_timer ? "ribbon: connect deferred to first idle"
                                : "ribbon: SetTimer failed; no ribbon buttons this session");
        if (!g_timer)
            FailSoft(L"A Windows timer could not be created to load the ribbon buttons.", "SetTimer failed");
    }

    void Stop()
    {
        g_stopping.store(true);
        core::SubscribeStateChanged(nullptr);
        if (g_timer) { KillTimer(nullptr, g_timer); g_timer = 0; }   // if it never fired

        // Disconnect, through the object Excel gave us; the window walk only if it gave none.
        if (g_connected.exchange(false))
        {
            bool sent = false;
            const bool hadKept = (g_addInInst != nullptr);
            if (hadKept) sent = PutBoolProp(g_addInInst, L"Connect", false);
            const bool sentByKept = sent;
            if (!sent)
            {
                std::ostringstream om;
                if (IDispatch* app = core::excelom::AcquireApplication(om))
                {
                    if (IDispatch* coll = GetDispProp(app, L"COMAddIns"))
                    {
                        if (IDispatch* entry = CallItem(coll, kProgId))
                        { sent = PutBoolProp(entry, L"Connect", false); entry->Release(); }
                        coll->Release();
                    }
                    app->Release();
                }
            }
            // which path disconnected, not just whether one did
            core::Log::Debug(!sent      ? "ribbon: disconnect did not take"
                             : sentByKept ? "ribbon: disconnected through the object Excel gave us"
                             : hadKept    ? "ribbon: kept object refused; disconnected by re-deriving"
                                          : "ribbon: no kept object; disconnected by re-deriving");
        }
        ReleaseRibbonUis();
        if (g_addInInst) { IDispatch* a = g_addInInst; g_addInInst = nullptr; a->Release(); }

        // only for a connect that failed and left the keys behind
        UnregisterServer();

        // Anything but 0 means Excel still holds the add-in object as the XLL goes.
        char line[96];
        _snprintf_s(line, _TRUNCATE, "ribbon: stopped; %ld add-in object(s) still alive",
                    g_objects.load());
        core::crashlog::Note(line);
        core::Log::Note(line);
    }

    bool Connected() { return g_connected.load() && !g_stopping.load(); }

    HRESULT GetClassObject(REFCLSID rclsid, REFIID riid, void** ppv)
    {
        if (rclsid != kClsid) return CLASS_E_CLASSNOTAVAILABLE;
        return g_factory.QueryInterface(riid, ppv);
    }
}
}
