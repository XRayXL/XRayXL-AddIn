// Unit test for the shared value decoder (vba/vbaretdecode.cpp) and the text grammar it is
// written in (core/textvalue.cpp), off Excel.
//
// The decoder's readers (core/safemem.h) read absolute addresses in this process under SEH, so
// the test needs no seam: it builds genuine COM structures with OleAut32, passes their real
// addresses to the decoder and asserts the rendering exactly.
//
// Built by XRayXL.sln into build\x64\Release\unit\, with vbaobject.cpp, which the decoder calls
// to describe an object.

#include "vbaretdecode.h"
#include "core/textvalue.h"
#include "core/jsonvalue.h"

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

    void Equal(const std::string& got, const char* want, const char* what)
    {
        const bool ok = got == want;
        printf("[%s] %s  (got: %s)\n", ok ? "PASS" : "FAIL", what, got.c_str());
        if (!ok) { printf("       want: %s\n", want); ++g_fail; }
    }

    // A VARIANT rendered through the text writer, into a buffer of `limit` bytes.
    std::string Render(const VARIANT& v, std::size_t limit = core::kMaxValueBytes)
    {
        core::TextBuf buf;
        buf.limit = limit;
        core::TextValueWriter w(buf);
        std::string s = vba::DescribeVariantValue(U(&v), w) ? buf.Text() : "<refused>";
        buf.Release();
        return s;
    }

    // The same VARIANT through the JSON writer. Each line is also printed as `JSON <text>` so
    // the driver can hand it to a real JSON parser.
    std::string RenderJson(const VARIANT& v)
    {
        core::TextBuf buf;
        core::JsonValueWriter w(buf);
        std::string s = vba::DescribeVariantValue(U(&v), w) ? buf.Text() : "<refused>";
        buf.Release();
        printf("JSON %s\n", s.c_str());
        return s;
    }

    SAFEARRAY* MakeVector(VARTYPE vt, LONG n, LONG lo = 0)
    {
        SAFEARRAYBOUND b{ static_cast<ULONG>(n), lo };
        return SafeArrayCreate(vt, 1, &b);
    }

    // A Variant array of `rows` by `cols`, both 1-based, as Range.Value2 hands one back.
    SAFEARRAY* MakeGrid(LONG rows, LONG cols)
    {
        SAFEARRAYBOUND b[2] = { { static_cast<ULONG>(rows), 1 }, { static_cast<ULONG>(cols), 1 } };
        return SafeArrayCreate(VT_VARIANT, 2, b);
    }

    void PutVar(SAFEARRAY* sa, LONG* idx, const VARIANT& v)
    {
        SafeArrayPutElement(sa, idx, const_cast<VARIANT*>(&v));
    }

    VARIANT ArrayVariant(VARTYPE vt, SAFEARRAY* sa)
    {
        VARIANT v; VariantInit(&v); v.vt = static_cast<VARTYPE>(VT_ARRAY | vt); v.parray = sa;
        return v;
    }
}

int main()
{
    // ---- scalars in a Variant: Double bare, every other number named ---------
    {
        VARIANT v; VariantInit(&v); v.vt = VT_I4; v.lVal = 42;
        Equal(Render(v), "Long(42)", "a Long in a Variant is named");
        v.vt = VT_R8; v.dblVal = 1.5;
        Equal(Render(v), "1.5", "a Double in a Variant is bare");
        v.vt = VT_BOOL; v.boolVal = VARIANT_TRUE;
        Equal(Render(v), "TRUE", "a Boolean in a Variant reads TRUE");
        v.vt = VT_DATE; v.date = 46352;
        Equal(Render(v), "Date(46352)", "a Date in a Variant is named, not read as a Double");
        v.vt = VT_CY; v.cyVal.int64 = 15000;
        Equal(Render(v), "Currency(1.5000)", "a Currency in a Variant is named and exact");
        v.vt = VT_EMPTY;
        Equal(Render(v), "Empty", "an Empty Variant");
    }

    // ---- a BSTR --------------------------------------------------------------
    {
        VARIANT v; VariantInit(&v); v.vt = VT_BSTR; v.bstrVal = SysAllocString(L"hello");
        Equal(Render(v), "\"hello\"", "a String in a Variant is quoted");
        VariantClear(&v);
    }

    // ---- a flat typed array: elements bare, both bounds ---------------------
    {
        SAFEARRAY* sa = MakeVector(VT_I4, 5);
        LONG* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        for (LONG i = 0; i < 5; ++i) data[i] = i;
        SafeArrayUnaccessData(sa);
        VARIANT v = ArrayVariant(VT_I4, sa);
        Equal(Render(v), "Long[0..4]{0,1,2,3,4}", "a typed array's elements are bare");
        VariantClear(&v);
    }

    // ---- an Enum array: VT_USERDEFINED, four bytes, Longs -------------------
    // VBA writes 29 in the vartype slot for `Dim e(1 To 2) As <Enum>`; the elements are
    // Longs and its own TypeName says Long(). The descriptor is built the same way here.
    {
        SAFEARRAY* sa = MakeVector(VT_I4, 3, 1);
        LONG* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        data[0] = 1; data[1] = 2; data[2] = 7;
        SafeArrayUnaccessData(sa);
        *(reinterpret_cast<DWORD*>(sa) - 1) = 29;      // VT_USERDEFINED, where SafeArrayGetVartype reads
        VARIANT v = ArrayVariant(VT_I4, sa);
        Equal(Render(v), "Long[1..3]{1,2,7}", "an Enum array reads as Longs");
        *(reinterpret_cast<DWORD*>(sa) - 1) = VT_I4;   // as OLE made it, so it frees as one
        VariantClear(&v);
    }

    // ---- a string array ------------------------------------------------------
    {
        SAFEARRAY* sa = MakeVector(VT_BSTR, 3, 1);
        BSTR* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        data[0] = SysAllocString(L"alpha");
        data[1] = SysAllocString(L"bravo");
        data[2] = SysAllocString(L"charlie");
        SafeArrayUnaccessData(sa);
        VARIANT v = ArrayVariant(VT_BSTR, sa);
        Equal(Render(v), "String[1..3]{\"alpha\",\"bravo\",\"charlie\"}", "a String array, Option Base 1");
        VariantClear(&v);
    }

    // ---- 2-D: one brace level per row, row 1 in full first -------------------
    {
        SAFEARRAY* sa = MakeGrid(2, 3);
        for (LONG r = 1; r <= 2; ++r)
            for (LONG c = 1; c <= 3; ++c)
            {
                VARIANT e; VariantInit(&e); e.vt = VT_R8; e.dblVal = r * 10 + c;
                LONG idx[2] = { r, c };
                PutVar(sa, idx, e);
            }
        VARIANT v = ArrayVariant(VT_VARIANT, sa);
        Equal(Render(v), "Variant[1..2,1..3]{{11,12,13},{21,22,23}}",
              "a 2-D array reads row by row, one level per row");
        VariantClear(&v);
    }

    // ---- a single row of a range is still 2-D --------------------------------
    {
        SAFEARRAY* sa = MakeGrid(1, 3);
        for (LONG c = 1; c <= 3; ++c)
        {
            VARIANT e; VariantInit(&e); e.vt = VT_R8; e.dblVal = c;
            LONG idx[2] = { 1, c };
            PutVar(sa, idx, e);
        }
        VARIANT v = ArrayVariant(VT_VARIANT, sa);
        Equal(Render(v), "Variant[1..1,1..3]{{1,2,3}}", "one row keeps its two levels");
        VariantClear(&v);
    }

    // ---- 3-D: the last index fastest, a level for each dimension -------------
    {
        SAFEARRAYBOUND b[3] = { { 2, 0 }, { 2, 0 }, { 2, 0 } };
        SAFEARRAY* sa = SafeArrayCreate(VT_I4, 3, b);
        for (LONG i = 0; i < 2; ++i)
            for (LONG j = 0; j < 2; ++j)
                for (LONG k = 0; k < 2; ++k)
                {
                    LONG idx[3] = { i, j, k };
                    LONG val = 100 * i + 10 * j + k;
                    SafeArrayPutElement(sa, idx, &val);
                }
        VARIANT v = ArrayVariant(VT_I4, sa);
        Equal(Render(v), "Long[0..1,0..1,0..1]{{{0,1},{10,11}},{{100,101},{110,111}}}",
              "a 3-D array nests three levels, last index fastest");
        VariantClear(&v);
    }

    // ---- mixed Variant elements: each names its type unless Double ----------
    {
        SAFEARRAY* sa = MakeVector(VT_VARIANT, 5);
        VARIANT* e = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&e));
        e[0].vt = VT_I2;   e[0].iVal = 1;
        e[1].vt = VT_R8;   e[1].dblVal = 2.5;
        e[2].vt = VT_BSTR; e[2].bstrVal = SysAllocString(L"x");
        e[3].vt = VT_BOOL; e[3].boolVal = VARIANT_FALSE;
        e[4].vt = VT_ERROR; e[4].scode = static_cast<SCODE>(0x800A07FA);   // #N/A
        SafeArrayUnaccessData(sa);
        VARIANT v = ArrayVariant(VT_VARIANT, sa);
        Equal(Render(v), "Variant[0..4]{Integer(1),2.5,\"x\",FALSE,#N/A}",
              "Variant elements: Integer named, Double bare, the rest self-typed");
        VariantClear(&v);
    }

    // ---- a Variant array of arrays: a nested array carries its own header ---
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
        VARIANT v = ArrayVariant(VT_VARIANT, outer);
        Equal(Render(v), "Variant[0..1]{Long[0..1]{10,20},Long[0..1]{30,40}}",
              "a nested array is an element with its own header, not a level");
        VariantClear(&v);   // frees the inner arrays too
    }

    // ---- an empty array: Array() is 0..-1 -----------------------------------
    {
        VARIANT v = ArrayVariant(VT_VARIANT, MakeVector(VT_VARIANT, 0));
        Equal(Render(v), "Variant[0..-1]{}", "an empty array states its bounds and has no elements");
        VariantClear(&v);
    }

    // ---- over the limit: the shape, never part of the contents --------------
    {
        SAFEARRAY* sa = MakeVector(VT_I4, 100);
        LONG* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        for (LONG i = 0; i < 100; ++i) data[i] = 1000000 + i;
        SafeArrayUnaccessData(sa);
        VARIANT v = ArrayVariant(VT_I4, sa);
        const long long before = core::ValuesRefused();
        Equal(Render(v, 64), "Long[0..99]", "an array over the limit keeps its header and no braces");
        Equal(core::ValuesRefused() == before + 1 ? "counted" : "not counted", "counted",
              "the refusal is counted for the log");
        VariantClear(&v);
    }

    // ---- JSON: every value names its type, Double included -------------------
    {
        VARIANT v; VariantInit(&v); v.vt = VT_R8; v.dblVal = 1.5;
        Equal(RenderJson(v), "{\"t\":\"Double\",\"v\":1.5}", "JSON names a Double");
        v.vt = VT_CY; v.cyVal.int64 = 15000;
        Equal(RenderJson(v), "{\"t\":\"Currency\",\"v\":\"1.5000\"}", "JSON keeps a Currency exact, as a string");
        v.vt = VT_BSTR; v.bstrVal = SysAllocString(L"a\"b\\c\n\x00e9");
        Equal(RenderJson(v), "{\"t\":\"String\",\"v\":\"a\\\"b\\\\c\\n\xc3\xa9\"}",
              "JSON escapes quotes, backslashes and newlines, and writes UTF-8");
        VariantClear(&v);
    }
    {
        SAFEARRAY* sa = MakeGrid(2, 2);
        const double vals[2][2] = { { 1, 2.5 }, { 3, 4 } };
        for (LONG r = 1; r <= 2; ++r)
            for (LONG c = 1; c <= 2; ++c)
            {
                VARIANT e; VariantInit(&e);
                if (r == 1 && c == 2) { e.vt = VT_BSTR; e.bstrVal = SysAllocString(L"x"); }
                else                  { e.vt = VT_R8; e.dblVal = vals[r - 1][c - 1]; }
                LONG idx[2] = { r, c };
                PutVar(sa, idx, e);
                VariantClear(&e);
            }
        VARIANT v = ArrayVariant(VT_VARIANT, sa);
        Equal(RenderJson(v),
              "{\"t\":\"Array\",\"elem\":\"Variant\",\"bounds\":[[1,2],[1,2]],\"v\":"
              "[[{\"t\":\"Double\",\"v\":1},{\"t\":\"String\",\"v\":\"x\"}],"
              "[{\"t\":\"Double\",\"v\":3},{\"t\":\"Double\",\"v\":4}]]}",
              "JSON: a Variant grid is rows of typed elements");
        VariantClear(&v);
    }
    {
        SAFEARRAY* sa = MakeVector(VT_I4, 3);
        LONG* data = nullptr; SafeArrayAccessData(sa, reinterpret_cast<void**>(&data));
        data[0] = 7; data[1] = 8; data[2] = 9;
        SafeArrayUnaccessData(sa);
        VARIANT v = ArrayVariant(VT_I4, sa);
        Equal(RenderJson(v), "{\"t\":\"Array\",\"elem\":\"Long\",\"bounds\":[[0,2]],\"v\":[7,8,9]}",
              "JSON: a typed array's elements are bare, typed by elem");
        VariantClear(&v);
    }
    {
        VARIANT v = ArrayVariant(VT_VARIANT, MakeVector(VT_VARIANT, 0));
        Equal(RenderJson(v), "{\"t\":\"Array\",\"elem\":\"Variant\",\"bounds\":[[0,-1]],\"v\":[]}",
              "JSON: an empty array");
        VariantClear(&v);
    }
    {
        core::TextBuf buf;
        core::JsonValueWriter w(buf);
        w.BeginArgs();
        w.BeginArg(1, "Long"); w.Number("Long", "5"); w.EndArg();
        w.BeginArg(2, "Variant"); w.BeginVariant(); w.Number("Integer", "3"); w.EndVariant(); w.EndArg();
        w.BeginArg(3, "Object"); w.BeginObject("Range", 0x10, "[B]S!A1"); w.Double(2); w.EndObject(); w.EndArg();
        w.BeginArg(4, "?unseen"); w.ArgUnreadable(); w.EndArg();
        w.ArgsNote(4, 6);
        w.EndArgs();
        printf("JSON %s\n", buf.Text());
        Equal(buf.Text(),
              "[{\"slot\":1,\"type\":\"Long\",\"value\":{\"t\":\"Long\",\"v\":5}},"
              "{\"slot\":2,\"type\":\"Variant\",\"value\":{\"t\":\"Integer\",\"v\":3}},"
              "{\"slot\":3,\"type\":\"Object\",\"value\":{\"t\":\"Object\",\"class\":\"Range\","
              "\"ptr\":\"0x10\",\"where\":\"[B]S!A1\",\"value\":{\"t\":\"Double\",\"v\":2}}},"
              "{\"slot\":4,\"type\":\"?unseen\",\"unreadable\":true},"
              "{\"described\":4,\"slots\":6}]",
              "JSON: an argument list");
        buf.Release();
    }

    // ---- a dynamic array never allocated: VT_ARRAY with a null SAFEARRAY ----
    {
        VARIANT v; VariantInit(&v);
        v.vt = VT_ARRAY | VT_I4;
        v.parray = nullptr;
        Equal(Render(v), "Long()", "an unallocated Long array names its type and has no bounds");
        Equal(RenderJson(v), "{\"t\":\"Array\",\"elem\":\"Long\",\"bounds\":[]}",
              "JSON: an unallocated array has no dimensions");
    }

    // ---- a bad address faults into a clean refusal, never a crash ------------
    {
        core::TextBuf buf;
        core::TextValueWriter w(buf);
        const bool ok = vba::DescribeBstrValue(0x40ull, w, true);  // unmapped
        Equal(ok ? "read" : "refused", "refused", "a bad address is refused, no crash");
        Equal(buf.Text(), "", "a refusal writes nothing");
        buf.Release();
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
