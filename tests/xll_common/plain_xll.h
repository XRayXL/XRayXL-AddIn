// The plain C-API boilerplate every ordinary XLL carries, shared by the test
// and demo add-ins. Not built from product sources, so a trace of these is
// a trace of a normal third-party XLL.
//
// Include it from exactly one .cpp per DLL: it defines Excel12 and the
// xlAutoClose / xlAutoFree12 exports.
#pragma once
#include <windows.h>
#include <cstdarg>
#include <cwchar>
#include "xlcall.h"

namespace plainxll
{
    typedef int(__stdcall* Excel12Proc)(int xlfn, int coper, LPXLOPER12* rgpxloper12,
                                        LPXLOPER12 xloper12Res);
    inline Excel12Proc g_pExcel12 = nullptr;

    inline void FetchEntryPoint()
    {
        if (g_pExcel12 != nullptr) return;
        HMODULE h = GetModuleHandleW(nullptr);
        if (h != nullptr)
            g_pExcel12 = reinterpret_cast<Excel12Proc>(GetProcAddress(h, "MdCallBack12"));
    }
}

// The SDK's own declaration in xlcall.h; this is its standard MdCallBack12 body.
extern "C" int __cdecl Excel12(int xlfn, LPXLOPER12 operRes, int count, ...)
{
    const int kMax = 32;
    LPXLOPER12 args[kMax];
    plainxll::FetchEntryPoint();
    if (plainxll::g_pExcel12 == nullptr) return xlretFailed;
    // Refuse rather than truncate: a silently shortened call is a wrong call.
    if (count < 0 || count > kMax) return xlretInvCount;
    va_list ap;
    va_start(ap, count);
    for (int i = 0; i < count; i++) args[i] = va_arg(ap, LPXLOPER12);
    va_end(ap);
    return plainxll::g_pExcel12(xlfn, count, args, operRes);
}

namespace plainxll
{
    // Filled in place rather than returned by value: the oper points into this
    // object's own buffer, so a copy would leave it aimed at the original.
    struct PascalStr
    {
        XCHAR buf[256];
        XLOPER12 oper;
        void Set(const wchar_t* text)
        {
            ZeroMemory(buf, sizeof(buf));
            ZeroMemory(&oper, sizeof(oper));
            size_t len = wcslen(text);
            if (len > 254) len = 254;
            buf[0] = static_cast<XCHAR>(len);
            for (size_t i = 0; i < len; i++) buf[i + 1] = static_cast<XCHAR>(text[i]);
            oper.xltype = xltypeStr;
            oper.val.str = buf;
        }
    };

    // Registers one worksheet function and returns xlfRegister's result code.
    // Stops at four arguments: the fifth is argument_text, not a category.
    inline int Register(XLOPER12& xDLL, const wchar_t* name, const wchar_t* typeText,
                        const wchar_t* displayName = nullptr)
    {
        PascalStr proc; proc.Set(name);
        PascalStr type; type.Set(typeText);
        PascalStr func; func.Set(displayName != nullptr ? displayName : name);
        XLOPER12 res; ZeroMemory(&res, sizeof(res));
        const int rc = Excel12(xlfRegister, &res, 4, &xDLL, &proc.oper, &type.oper, &func.oper);
        Excel12(xlFree, nullptr, 1, &res);
        return rc;
    }

    // Per-thread, because an XLL function may run on any calculation worker.
    inline thread_local XLOPER12 t_ret;
    inline thread_local XCHAR    t_str[512];

    inline LPXLOPER12 RetNum(double v)
    {
        t_ret.xltype = xltypeNum; t_ret.val.num = v; return &t_ret;
    }
    inline LPXLOPER12 RetStr(const wchar_t* s)
    {
        size_t len = wcslen(s); if (len > 500) len = 500;
        t_str[0] = static_cast<XCHAR>(len);
        for (size_t i = 0; i < len; i++) t_str[i + 1] = static_cast<XCHAR>(s[i]);
        t_ret.xltype = xltypeStr; t_ret.val.str = t_str; return &t_ret;
    }
    inline LPXLOPER12 RetErr(int e)
    {
        t_ret.xltype = xltypeErr; t_ret.val.err = e; return &t_ret;
    }

    // A string allocated per call and marked xlbitDLLFree: Excel copies it, then passes it to
    // xlAutoFree12. One HeapAlloc block with the XLOPER12 first, which is what xlAutoFree12 frees.
    struct OwnedStr { XLOPER12 x; XCHAR s[256]; };
    inline LPXLOPER12 RetOwnedStr(const wchar_t* text)
    {
        OwnedStr* o = static_cast<OwnedStr*>(HeapAlloc(GetProcessHeap(), 0, sizeof(OwnedStr)));
        if (o == nullptr) return RetErr(xlerrNum);
        size_t len = wcslen(text); if (len > 255) len = 255;
        o->s[0] = static_cast<XCHAR>(len);
        for (size_t i = 0; i < len; i++) o->s[i + 1] = static_cast<XCHAR>(text[i]);
        o->x.xltype = xltypeStr | xlbitDLLFree;
        o->x.val.str = o->s;
        return &o->x;
    }
}

extern "C" __declspec(dllexport) int  __stdcall xlAutoClose() { return 1; }
// Frees what RetOwnedStr returned; Excel calls it for any result marked xlbitDLLFree.
extern "C" __declspec(dllexport) void __stdcall xlAutoFree12(LPXLOPER12 p)
{
    if (p != nullptr && (p->xltype & xlbitDLLFree)) HeapFree(GetProcessHeap(), 0, p);
}
