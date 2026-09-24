// Unit test for the trace-param parsers (app/paramparse.cpp), off Excel. The parsers make no
// Excel call, so the test builds XLOPER12s by hand and asserts the grammar: what is accepted,
// what is refused, and the BUFFERSIZE units.
//
// Built by XRayXL.sln into build\x64\Release\unit\.

#include "paramparse.h"

#include <cstdio>
#include <cstring>
#include <clocale>

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    // Build a string XLOPER12 into a caller-owned pascal buffer.
    XLOPER12 Str(XCHAR* buf, const wchar_t* s)
    {
        int n = 0; while (s[n]) ++n;
        buf[0] = static_cast<XCHAR>(n);
        for (int i = 0; i < n; ++i) buf[i + 1] = static_cast<XCHAR>(s[i]);
        XLOPER12 x{}; x.xltype = xltypeStr; x.val.str = buf; return x;
    }
    XLOPER12 Int(int n)     { XLOPER12 x{}; x.xltype = xltypeInt;  x.val.w = n; return x; }
    XLOPER12 Num(double d)  { XLOPER12 x{}; x.xltype = xltypeNum;  x.val.num = d; return x; }
    XLOPER12 Bool(bool b)   { XLOPER12 x{}; x.xltype = xltypeBool; x.val.xbool = b; return x; }
    XLOPER12 Missing()      { XLOPER12 x{}; x.xltype = xltypeMissing; return x; }

    unsigned long long ParseBuf(XLOPER12& x)
    {
        unsigned long long b = ~0ull; return app::params::ParseBufferBytes(&x, b) ? b : ~0ull;
    }
}

int main()
{
    using namespace app::params;
    XCHAR s0[64], s1[64];

    // ---- Source: XLL | VBA | (omitted = both), case-insensitive -------------
    {
        bool both = false; core::modes::Source s{};
        auto xll = Str(s0, L"xll");   // lower-case on purpose
        Check(ParseSource(&xll, both, s) && !both && s == core::modes::Source::Xll, "Source xll -> XLL");
        auto vba = Str(s0, L"VBA");
        Check(ParseSource(&vba, both, s) && !both && s == core::modes::Source::Vba, "Source VBA -> VBA");
        auto miss = Missing();
        Check(ParseSource(&miss, both, s) && both, "Source omitted -> both");
        auto com = Str(s0, L"COM");
        Check(!ParseSource(&com, both, s), "Source COM refused");
    }

    // ---- Param, Depth, OnOff, WhenFull --------------------------------------
    {
        core::modes::Param p{};
        auto d = Str(s0, L"DEPTH");  Check(ParseParam(&d, p) && p == core::modes::Param::Depth, "Param DEPTH");
        auto a = Str(s0, L"ARGS");   Check(ParseParam(&a, p) && p == core::modes::Param::Args,  "Param ARGS");
        auto j = Str(s0, L"JUNK");   Check(!ParseParam(&j, p), "Param JUNK refused");

        core::modes::Depth dp{};
        auto off = Str(s0, L"off");  Check(ParseDepth(&off, dp) && dp == core::modes::Depth::Off, "Depth off -> OFF");
        auto all = Str(s0, L"ALL");  Check(ParseDepth(&all, dp) && dp == core::modes::Depth::All, "Depth ALL");
        auto ban = Str(s0, L"BANANAS"); Check(!ParseDepth(&ban, dp), "Depth BANANAS refused");

        bool on = false;
        auto bt = Bool(true);        Check(ParseOnOff(&bt, on) && on, "OnOff native TRUE");
        auto tf = Str(s0, L"FALSE"); Check(ParseOnOff(&tf, on) && !on, "OnOff text FALSE");
        auto mb = Str(s0, L"MAYBE"); Check(!ParseOnOff(&mb, on), "OnOff MAYBE refused");
        auto iv = Int(1);            Check(!ParseOnOff(&iv, on), "OnOff int refused");

        bool pause = true;
        auto dr = Str(s0, L"DROP");  Check(ParseWhenFull(&dr, pause) && !pause, "WhenFull DROP");
        auto pa = Str(s0, L"pause"); Check(ParseWhenFull(&pa, pause) && pause, "WhenFull pause -> PAUSE");
        auto wx = Str(s0, L"X");     Check(!ParseWhenFull(&wx, pause), "WhenFull X refused");
    }

    // ---- BUFFERSIZE units: bare/M/MB = MB, K/KB = KB, case-insensitive ----------
    {
        const unsigned long long MB = 1024ull * 1024ull, KB = 1024ull;
        auto b64  = Str(s0, L"64");     Check(ParseBuf(b64)  == 64 * MB,  "BUFFERSIZE 64 -> 64 MB");
        auto b64m = Str(s0, L"64MB");   Check(ParseBuf(b64m) == 64 * MB,  "BUFFERSIZE 64MB -> 64 MB");
        auto b64l = Str(s0, L"64mb");   Check(ParseBuf(b64l) == 64 * MB,  "BUFFERSIZE 64mb (case) -> 64 MB");
        auto b512k= Str(s0, L"512K");   Check(ParseBuf(b512k)== 512 * KB, "BUFFERSIZE 512K -> 512 KB");
        auto b512b= Str(s0, L"512KB");  Check(ParseBuf(b512b)== 512 * KB, "BUFFERSIZE 512KB -> 512 KB");
        auto bsp  = Str(s0, L"512 KB"); Check(ParseBuf(bsp)  == 512 * KB, "BUFFERSIZE '512 KB' (space) -> 512 KB");
        auto bint = Int(64);            Check(ParseBuf(bint) == 64 * MB,  "BUFFERSIZE int 64 -> 64 MB");
        auto bnum = Num(32);            Check(ParseBuf(bnum) == 32 * MB,  "BUFFERSIZE num 32 -> 32 MB");
        auto b0   = Str(s0, L"0");      Check(ParseBuf(b0)   == 0,        "BUFFERSIZE 0 -> synchronous");
        auto bj   = Str(s0, L"junk");   Check(ParseBuf(bj)   == ~0ull,    "BUFFERSIZE junk refused");
        auto bdot = Str(s0, L".");      Check(ParseBuf(bdot) == ~0ull,    "BUFFERSIZE '.' refused: a dot is not a number");
        auto bdk  = Str(s0, L".KB");    Check(ParseBuf(bdk)  == ~0ull,    "BUFFERSIZE '.KB' refused");
        auto bmax = Str(s0, L"240MB");  Check(ParseBuf(bmax) == 240 * MB, "BUFFERSIZE 240MB is the most accepted");
        auto bover = Str(s0, L"241MB"); Check(ParseBuf(bover) == ~0ull,   "BUFFERSIZE over 240 MB refused");
        auto bbig = Str(s0, L"9999MB"); Check(ParseBuf(bbig) == ~0ull,    "BUFFERSIZE far over the cap refused");
    }

    // ---- FormatBufferW round-trips through ParseBufferBytes ------------------
    {
        wchar_t out[32];
        FormatBufferW(64ull * 1024 * 1024, out, 32); Check(!wcscmp(out, L"64MB"), "Format 64 MB -> 64MB");
        FormatBufferW(512ull * 1024, out, 32);       Check(!wcscmp(out, L"512KB"), "Format 512 KB -> 512KB");
        FormatBufferW(0, out, 32);                   Check(!wcscmp(out, L"0"), "Format 0 -> 0");
        // round-trip: format 48 MB, parse it back
        FormatBufferW(48ull * 1024 * 1024, out, 32);
        auto rt = Str(s0, out);
        Check(ParseBuf(rt) == 48ull * 1024 * 1024, "Format then Parse round-trips (48 MB)");
    }

    // ---- IsWord is case-insensitive -----------------------------------------
    {
        // The full word in the wrong case. A prefix must not match, or "buffer" would name
        // BUFFERSIZE and BUFFERWHENFULL alike.
        auto w = Str(s0, L"buffersize");
        Check(IsWord(&w, L"BUFFERSIZE"), "IsWord matches case-insensitively");
        auto wp = Str(s0, L"buffer");
        Check(!IsWord(&wp, L"BUFFERSIZE"), "IsWord does not match a prefix");
        auto w2 = Str(s1, L"BUFFERWHENFULL");
        Check(!IsWord(&w2, L"BUFFERSIZE"), "IsWord rejects a different word");
    }

    // ---- a fractional size, under a comma-decimal locale ---------------------
    // Parsing must not depend on the locale's decimal separator.
    {
        XCHAR loc[64];
        unsigned long long b = 0;

        auto onePointFive = Str(loc, L"1.5");
        Check(ParseBufferBytes(&onePointFive, b) && b == 1572864ull,
              "BUFFERSIZE 1.5 -> 1.5 MB in the default locale");

        const char* had = setlocale(LC_NUMERIC, "de-DE");
        if (had == nullptr)
        {
            printf("[SKIP] comma-decimal locale unavailable on this machine\n");
        }
        else
        {
            b = 0;
            auto again = Str(loc, L"1.5");
            Check(ParseBufferBytes(&again, b) && b == 1572864ull,
                  "BUFFERSIZE 1.5 -> 1.5 MB under a comma-decimal locale");

            b = 0;
            auto k = Str(loc, L"1.5KB");
            Check(ParseBufferBytes(&k, b) && b == 1536ull,
                  "BUFFERSIZE 1.5KB -> 1536 bytes under a comma-decimal locale");
            setlocale(LC_NUMERIC, "C");
        }
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
