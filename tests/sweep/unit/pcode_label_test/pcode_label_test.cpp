// Unit test for the Variant label in vba::ReadArgTypes (vba/vbapcode.cpp): a parameter the body
// only passes on is typed by the CVarRef or CDargRef right after its push, whose VARTYPE is what
// the callee reads it by. A typed instruction anywhere in the body still decides.
//
// Each case is one procedure in memory: a last statement, the instructions, exit 635, the trailer.
//
// Built by XRayXL.sln into build\x64\Release\unit\; the log and the corpus's module walk are
// stubbed below.

#include "vbapcode.h"
#include "vbatrailer.h"

#include <cstdio>
#include <cstring>
#include <initializer_list>
#include <string>
#include <vector>

namespace core { namespace Log { void Warning(const std::string&) {} } }
namespace vba { int ModuleTrailers(std::uint64_t, std::uint64_t*, int) { return 0; } }   // the corpus is not tested here

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    constexpr std::uint16_t kBos = 615, kExit = 635, kFiller = 1100;
    constexpr std::uint16_t kFLdRf = 671, kFLdAd = 751, kLoadLong = 658;
    constexpr std::uint16_t kCDargRef = 950, kCVarRef = 951, kRedim = 1473, kRedimPreserve = 1474;

    struct Code
    {
        std::vector<std::uint8_t> b;
        Code& W(std::uint16_t v) { b.push_back(static_cast<std::uint8_t>(v)); b.push_back(static_cast<std::uint8_t>(v >> 8)); return *this; }
        Code& D(std::int32_t v) { for (int i = 0; i < 4; ++i) b.push_back(static_cast<std::uint8_t>(v >> (8 * i))); return *this; }
        Code& Push(std::uint16_t op, int slot) { return W(op).D(8 * slot); }
        Code& VarRef(std::uint16_t vt) { return W(kCVarRef).D(-32).W(vt); }
        Code& DargRef(std::uint16_t vt) { return W(kCDargRef).W(vt); }
        Code& Filler() { return W(kFiller); }
        Code& LoadLong(int slot) { return W(kLoadLong).D(8 * slot); }
        // One dimension of `elem`, 8-byte elements, FADF_HAVEVARTYPE: a Double array's real operand.
        Code& Redim(std::uint16_t op, std::uint16_t elem) { return W(op).W(1).W(elem).W(8).W(0x80); }
    };

    alignas(8) std::uint8_t g_mem[512];

    // The body follows a last statement (BoS with no successor) and ends in the exit.
    bool Walk(const Code& body, vba::ArgTypes& out)
    {
        Code c;
        c.W(kBos).D(0);
        c.b.insert(c.b.end(), body.b.begin(), body.b.end());
        c.W(kExit);
        std::memset(g_mem, 0, sizeof g_mem);
        std::memcpy(g_mem + 64, c.b.data(), c.b.size());
        const std::uint16_t procSize = static_cast<std::uint16_t>(c.b.size());
        std::uint8_t* trailer = g_mem + 64 + procSize;
        std::memcpy(trailer + vba::kTrl_procSize, &procSize, 2);
        return vba::ReadArgTypes(reinterpret_cast<std::uint64_t>(trailer), out) && !out.partial;
    }

    bool Is(const char* got, const char* want) { return got && want ? strcmp(got, want) == 0 : got == want; }

    // As an entry capture does: walk, then count the labels.
    void Typed(const char* what, const Code& body, int slot, const char* want)
    {
        vba::ArgTypes out;
        const bool walked = Walk(body, out);
        printf("  slot %d: %s\n", slot, out.name[slot] ? out.name[slot] : "(untyped)");
        Check(walked && Is(out.name[slot], want), what);
        vba::NoteLabels(out);
    }
}

int main()
{
    static vba::PcodeLengths L;
    L.slots = 1700; L.pinned = 1; L.ok = true;
    for (std::uint16_t op : { kBos, kFLdRf, kFLdAd, kLoadLong }) L.len[op] = 6;
    L.len[kCVarRef] = 8; L.len[kCDargRef] = 4; L.len[kFiller] = 2;
    L.len[kRedim] = 10; L.len[kRedimPreserve] = 10;
    vba::SetArmedLengths(L);
    vba::ResetPcodeCounts();

    // Demo_AddMonths(ByVal Dte As Date) handing Dte to Year(): these are its bytes.
    {
        Code c;
        const std::uint8_t real[] = { 0x9F, 0x02, 0x08, 0, 0, 0, 0xB7, 0x03, 0xE0, 0xFF, 0xFF, 0xFF, 0x07, 0x40 };
        c.b.assign(real, real + sizeof real);
        Typed("a ByVal Date handed to a Variant reads Double, as its load would name it", c, 1, "Double");
    }
    Typed("a ByRef String through CDargRef reads String&", Code().Push(kFLdAd, 2).DargRef(0x4008), 2, "String&");
    Typed("a ByRef Boolean reads Integer&", Code().Push(kFLdAd, 1).DargRef(0x400B), 1, "Integer&");
    Typed("a ByVal Long reads Long", Code().Push(kFLdRf, 1).VarRef(0x4003), 1, "Long");
    Typed("a ByRef Object reads Object&", Code().Push(kFLdAd, 1).VarRef(0x4009), 1, "Object&");
    Typed("a ByRef Variant reads Variant&", Code().Push(kFLdAd, 1).VarRef(0x400C), 1, "Variant&");
    Typed("a ByRef array reads Ref&, as 747 names it", Code().Push(kFLdAd, 1).VarRef(0x6005), 1, "Ref&");
    Typed("a ByRef LongLong reads Ref&, as 747 names it", Code().Push(kFLdAd, 1).VarRef(0x4014), 1, "Ref&");
    Typed("a ReDim of an array parameter reads Ref&", Code().Push(kFLdAd, 2).Redim(kRedim, 5), 2, "Ref&");
    Typed("...and so does a ReDim Preserve", Code().Push(kFLdAd, 1).Redim(kRedimPreserve, 8), 1, "Ref&");
    Typed("a ReDim one instruction late types nothing",
          Code().Push(kFLdAd, 1).Filler().Redim(kRedim, 5), 1, nullptr);

    Typed("a label one instruction late types nothing",
          Code().Push(kFLdRf, 1).Filler().VarRef(0x4007), 1, nullptr);
    Typed("the label types the slot pushed last, not an earlier one",
          Code().Push(kFLdRf, 1).Push(kFLdRf, 3).VarRef(0x4008), 3, "String");
    Typed("...and leaves the earlier one untyped",
          Code().Push(kFLdRf, 1).Push(kFLdRf, 3).VarRef(0x4008), 1, nullptr);
    Typed("a typed load earlier in the body decides",
          Code().LoadLong(1).Push(kFLdRf, 1).VarRef(0x4005), 1, "Long");
    Typed("a typed load later in the body decides",
          Code().Push(kFLdRf, 1).VarRef(0x4005).LoadLong(1), 1, "Long");

    vba::ResetPcodeCounts();
    Typed("a ByVal Variant label types nothing: that is three slots", Code().Push(kFLdRf, 1).VarRef(0x400C), 1, nullptr);
    Typed("a VARTYPE without VT_BYREF types nothing", Code().Push(kFLdRf, 1).VarRef(0x0003), 1, nullptr);
    Typed("a ByVal array label types nothing", Code().Push(kFLdRf, 1).VarRef(0x6005), 1, nullptr);
    Typed("an array of an element type not named types nothing", Code().Push(kFLdAd, 1).VarRef(0x608A), 1, nullptr);
    Typed("a ReDim of an element type not named types nothing", Code().Push(kFLdAd, 1).Redim(kRedim, 0x8A), 1, nullptr);
    Typed("a ReDim after a ByVal push types nothing", Code().Push(kFLdRf, 1).Redim(kRedim, 5), 1, nullptr);
    const std::string line = vba::PcodeLine();
    printf("  %s\n", line.c_str());
    Check(line.find("0 parameter(s) typed by a Variant label (6 label(s) fit no type)") != std::string::npos,
          "the disarm line counts the labels that fit no type");

    // The exit re-read walks the procedure again and must not count it twice.
    vba::ResetPcodeCounts();
    {
        vba::ArgTypes out;
        Walk(Code().Push(kFLdRf, 1).VarRef(0x4003), out);
        vba::NoteLabels(out);
        Walk(Code().Push(kFLdRf, 1).VarRef(0x4003), out);
        const std::string counted = vba::PcodeLine();
        Check(out.labelled == 1 && counted.find(", 1 parameter(s) typed by a Variant label") != std::string::npos,
              "a walk alone adds nothing to the disarm line");
    }

    // A build whose table disagrees on CVarRef's length leaves the operand unread.
    L.len[kCVarRef] = 6;
    vba::SetArmedLengths(L);
    Typed("a CVarRef of another length types nothing", Code().Push(kFLdRf, 1).W(kCVarRef).D(-32), 1, nullptr);

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
