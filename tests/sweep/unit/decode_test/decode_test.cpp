// UNIT TEST for the shared value decoder (vba/vbaretdecode.cpp), off Excel.
//
// The decoder turns Excel/COM VARIANT/SAFEARRAY/BSTR bytes into trace text, and
// its readers (core/safemem.h) read absolute addresses in THIS process under SEH.
// So the test needs no seam and no Excel: build GENUINE COM structures with
// OleAut32 (the same layout Excel produces), pass their real addresses to the
// decoder, and assert the rendering -- including the nested-array render-in-place
// that a per-level buffer used to cut at 512 bytes.
//
// Built by XRayXL.sln into build\x64\Release\unit\, with vbaobject.cpp, which
// the decoder calls to describe an object.

#include "vbaretdecode.h"

#include <windows.h>
#include <oleauto.h>
#include <cstdio>
#include <cstdint>
#include <cstring>
#include <string>

namespace
{
    int g_fail = 0;
    std::uint64_t U(const void* p) { return reinterpret_cast<std::uint64_t>(p); }

    void Contains(const char* got, const char* needle, const char* what)
    {
        const bool ok = std::strstr(got, needle) != nullptr;
        printf("[%s] %s  (got: \"%s\")\n", ok ? "PASS" : "FAIL", what, got);
        if (!ok) ++g_fail;
    }
    void Absent(const char* got, const char* needle, const char* what)
    {
        const bool ok = std::strstr(got, needle) == nullptr;
        printf("[%s] %s  (got: \"%s\")\n", ok ? "PASS" : "FAIL", what, got);
        if (!ok) ++g_fail;
    }

    // A one-dimensional SAFEARRAY of `vt`, lower bound 0, `n` elements.
    SAFEARRAY* MakeVector(VARTYPE vt, LONG n)
    {
        SAFEARRAYBOUND b{ static_cast<ULONG>(n), 0 };
        return SafeArrayCreate(vt, 1, &b);
    }
}

int main()
{
    static char buf[64 * 1024];
    const char* held = "";

    // ---- a scalar Variant (Long 42) -----------------------------------------
    {
        VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = 42;
        buf[0] = 0;
        vba::DescribeVariantValue(U(&v), buf, sizeof buf, &held);
        Contains(buf, "42", "scalar Long variant renders its value");
        VariantClear(&v);
    }

    // ---- a BSTR --------------------------------------------------------------
    {
        BSTR s = SysAllocString(L"hello");
        buf[0] = 0;
        vba::DescribeBstrValue(U(s), buf, sizeof buf, /*told=*/true);
        Contains(buf, "hello", "BSTR renders its text");
        SysFreeString(s);
    }

    // ---- a flat array: Long[0..4]{0,1,2,3,4} ---------------------------------
    {
        SAFEARRAY* sa = MakeVector(VT_I4, 5);
        LONG* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        for (LONG i = 0; i < 5; ++i) data[i] = i;
        SafeArrayUnaccessData(sa);
        VARIANT v; VariantInit(&v); v.vt = VT_ARRAY | VT_I4; v.parray = sa;
        buf[0] = 0;
        vba::DescribeVariantValue(U(&v), buf, sizeof buf, &held);
        Contains(buf, "0,1,2,3,4", "flat Long array renders all elements");
        VariantClear(&v);
    }

    // ---- a long string array: elements are BSTRs, not cut at a small temp ----
    {
        SAFEARRAY* sa = MakeVector(VT_BSTR, 3);
        BSTR* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        data[0] = SysAllocString(L"alpha");
        data[1] = SysAllocString(L"bravo");
        data[2] = SysAllocString(L"charlie");
        SafeArrayUnaccessData(sa);
        VARIANT v; VariantInit(&v); v.vt = VT_ARRAY | VT_BSTR; v.parray = sa;
        buf[0] = 0;
        vba::DescribeVariantValue(U(&v), buf, sizeof buf, &held);
        Contains(buf, "alpha", "string array renders first element");
        Contains(buf, "charlie", "string array renders last element");
        VariantClear(&v);
    }

    // ---- a NESTED array: Variant[]{ Long[]{10,20}, Long[]{30,40} } -----------
    // The render-in-place decoder must render the inner arrays IN FULL,
    // not cut them to "[...]" as the old fixed per-level 512-byte temp would.
    {
        SAFEARRAY* outer = MakeVector(VT_VARIANT, 2);
        VARIANT* ov = nullptr; SafeArrayAccessData(outer, reinterpret_cast<void**>(&ov));
        for (int k = 0; k < 2; ++k)
        {
            SAFEARRAY* inner = MakeVector(VT_I4, 2);
            LONG* id = nullptr; SafeArrayAccessData(inner, reinterpret_cast<void**>(&id));
            id[0] = (k == 0) ? 10 : 30;
            id[1] = (k == 0) ? 20 : 40;
            SafeArrayUnaccessData(inner);
            VariantInit(&ov[k]); ov[k].vt = VT_ARRAY | VT_I4; ov[k].parray = inner;
        }
        SafeArrayUnaccessData(outer);
        VARIANT v; VariantInit(&v); v.vt = VT_ARRAY | VT_VARIANT; v.parray = outer;
        buf[0] = 0;
        vba::DescribeVariantValue(U(&v), buf, sizeof buf, &held);
        Contains(buf, "10,20", "nested array renders inner elements 10,20 (render-in-place)");
        Contains(buf, "30,40", "nested array renders inner elements 30,40");
        Absent(buf, "[...]", "nested array is NOT cut off with [...]");
        VariantClear(&v);   // frees the inner arrays too
    }

    // ---- a bad address faults into a clean refusal, never a crash ------------
    {
        buf[0] = 0;
        const bool ok = vba::DescribeBstrValue(0x40ull, buf, sizeof buf, true);  // unmapped
        printf("[%s] bad address refused, no crash (returned %d)\n", ok ? "FAIL" : "PASS", ok);
        if (ok) ++g_fail;
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
