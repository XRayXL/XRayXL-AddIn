// Unit test for xll::Parse (xll/xlltypeplan.cpp): '%' widens only the codes the
// C API gives a wide form -- C D F G K O. After any other code the registration
// is malformed, and must be declined rather than traced as if the '%' were absent.
//
// Built by XRayXL.sln into build\x64\Release\unit\.

#include "xlltypeplan.h"

#include <cstdio>

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    void Accepts(const wchar_t* t, const char* label) { Check(xll::Parse(t).ok, label); }
    void Refuses(const wchar_t* t, const char* label) { Check(!xll::Parse(t).ok, label); }
}

int main()
{
    // ---- the wide forms the C API defines ------------------------------------
    Accepts(L"QC%",  "argument C% is accepted");
    Accepts(L"QD%",  "argument D% is accepted");
    Accepts(L"QF%",  "argument F% is accepted");
    Accepts(L"QG%",  "argument G% is accepted");
    Accepts(L"QK%",  "argument K% is accepted");
    Accepts(L"QO%",  "argument O% is accepted");
    Accepts(L"C%B",  "return C% is accepted");
    Accepts(L"D%B",  "return D% is accepted");
    Accepts(L"K%B",  "return K% is accepted");
    Accepts(L"QBB$", "a modifier after plain codes is accepted");

    // ---- '%' after a code with no wide form ----------------------------------
    Refuses(L"QB%",  "argument B% is refused");
    Refuses(L"QE%",  "argument E% is refused");
    Refuses(L"QJ%",  "argument J% is refused");
    Refuses(L"QP%",  "argument P% is refused");
    Refuses(L"QQ%",  "argument Q% is refused");
    Refuses(L"B%B",  "return B% is refused");
    Refuses(L"Q%B",  "return Q% is refused");

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
