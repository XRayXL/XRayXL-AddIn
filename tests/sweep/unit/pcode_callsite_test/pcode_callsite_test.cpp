// Unit test for reading a caller's call site (vba/vbapcode.cpp): the pushes straight before a
// call, each one operand-stack slot, the first pushed first; a local's type from the typed
// instructions on its frame slot; and the type each literal slot carries.
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

    constexpr std::uint16_t kBos = 615, kExit = 635, kCall = 1311;
    constexpr std::uint16_t kLitI4 = 1520, kLitStr = 1527, kLitR8 = 1524, kMulI4 = 244;
    constexpr std::uint16_t kFLdRf = 671, kFLdAd = 751, kFLdI4 = 658, kFLdFPR8 = 669;
    constexpr std::uint16_t kFStI4 = 690, kFStStrCopy = 708, kFStVar = 694, kRedim = 1473;

    struct Code
    {
        std::vector<std::uint8_t> b;
        Code& W(std::uint16_t v) { b.push_back(static_cast<std::uint8_t>(v)); b.push_back(static_cast<std::uint8_t>(v >> 8)); return *this; }
        Code& D(std::int32_t v) { for (int i = 0; i < 4; ++i) b.push_back(static_cast<std::uint8_t>(v >> (8 * i))); return *this; }
        Code& Op(std::uint16_t op, std::int32_t operand) { return W(op).D(operand); }
        Code& Str() { return W(kLitStr).W(0); }
        Code& Call(std::uint16_t argBytes, std::uint16_t index = 0) { return W(kCall).W(index).W(argBytes); }
        Code& Redim(std::uint16_t elem) { return W(kRedim).W(1).W(elem).W(8).W(0x80); }
    };

    alignas(8) std::uint8_t g_mem[512];
    std::uint32_t g_callOff = 0;

    // The body follows a last statement and ends in the exit; returns the trailer.
    std::uint64_t Place(const Code& body)
    {
        Code c;
        c.Op(kBos, 0);
        const std::size_t bodyAt = c.b.size();
        c.b.insert(c.b.end(), body.b.begin(), body.b.end());
        c.W(kExit);
        g_callOff = 0;
        for (std::size_t i = bodyAt; i + 1 < c.b.size(); ++i)
            if (c.b[i] == (kCall & 0xFF) && c.b[i + 1] == (kCall >> 8)) { g_callOff = static_cast<std::uint32_t>(i); break; }
        std::memset(g_mem, 0, sizeof g_mem);
        std::memcpy(g_mem + 64, c.b.data(), c.b.size());
        const std::uint16_t procSize = static_cast<std::uint16_t>(c.b.size());
        std::uint8_t* trailer = g_mem + 64 + procSize;
        std::memcpy(trailer + vba::kTrl_procSize, &procSize, 2);
        return reinterpret_cast<std::uint64_t>(trailer);
    }

    bool Is(const char* got, const char* want) { return got && want ? strcmp(got, want) == 0 : got == want; }
}

int main()
{
    static vba::PcodeLengths L;
    L.slots = 1700; L.pinned = 1; L.ok = true;
    for (std::uint16_t op : { kBos, kLitI4, kFLdRf, kFLdAd, kFLdI4, kFLdFPR8, kCall, kFStI4, kFStStrCopy, kFStVar }) L.len[op] = 6;
    L.len[kLitStr] = 4; L.len[kLitR8] = 10; L.len[kMulI4] = 2; L.len[kRedim] = 10;
    vba::SetArmedLengths(L);
    vba::PcodePush p[4] = {};

    // `Callee 42, "hi", n`: pushed last to first, so the local's address is pushed first.
    {
        const std::uint64_t t = Place(Code().Op(kFLdRf, -16).Str().Op(kLitI4, 42).Call(0x18));
        const bool ok = vba::ReadCallPushes(t, g_callOff, 3, p);
        Check(ok && p[0].kind == vba::PushKind::SlotAddress && p[0].operand == -16 &&
              p[1].kind == vba::PushKind::Literal && Is(p[1].type, "String") &&
              p[2].kind == vba::PushKind::Literal && Is(p[2].type, "Long"),
              "three pushes read in order: an address, a string, a long");
    }
    {
        const std::uint64_t t = Place(Code().Op(kFLdI4, -8).Op(kFLdAd, 8).Call(0x10));
        Check(vba::ReadCallPushes(t, g_callOff, 2, p) && p[0].kind == vba::PushKind::Value &&
              Is(p[0].type, "Long") && p[1].kind == vba::PushKind::HeldPointer && p[1].operand == 8,
              "a typed load is a value, a ByRef parameter's pointer is held");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 1).Op(kLitI4, 2).Call(0x08));
        Check(vba::ReadCallPushes(t, g_callOff, 1, p) && p[0].operand == 2,
              "only the last n instructions are the call's: `x = a + F(b)` pushes a first");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 12).Op(kFLdI4, -16).W(kMulI4).Call(0x08));
        Check(!vba::ReadCallPushes(t, g_callOff, 1, p), "an expression is declined");
    }
    {
        const std::uint64_t t = Place(Code().Op(kFLdFPR8, -8).Call(0x08));
        Check(!vba::ReadCallPushes(t, g_callOff, 1, p), "an FPU load is not a push");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 3).Call(0x10));
        Check(!vba::ReadCallPushes(t, g_callOff, 2, p), "fewer instructions than slots is declined");
    }

    // A local's type.
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 7).Op(kFStI4, -16).Op(kFLdRf, -16).Call(0x08));
        Check(Is(vba::ReadLocalTypeName(t, -16), "Long"), "a local stored by FStI4 is a Long");
        Check(Is(vba::ReadLocalTypeName(t, -24), nullptr), "a local nothing touches has no type");
    }
    {
        const std::uint64_t t = Place(Code().Str().Op(kFStStrCopy, -24).Op(kLitI4, 5).Op(kFStVar, -48));
        Check(Is(vba::ReadLocalTypeName(t, -24), "String"), "a local stored by FStStrCopy is a String");
        Check(Is(vba::ReadLocalTypeName(t, -48), "Variant"), "a local stored by FStVar is a Variant");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 2).Op(kFLdRf, -64).Redim(5));
        Check(Is(vba::ReadLocalTypeName(t, -64), "Ref"), "a local ReDimmed is an array");
    }

    // The calls a procedure passes one of its parameters to.
    {
        vba::CallPass c[4] = {};
        std::uint64_t t = Place(Code().Op(kFLdAd, 8).Call(0x08, 5));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 1 && c[0].index == 5 && c[0].argBytes == 8 &&
              c[0].argSlot == 1 && c[0].pushOp == kFLdAd,
              "a ByRef parameter passed on is found, with the pool index and the slot it lands in");
        t = Place(Code().Op(kFLdRf, 16).Op(kLitI4, 3).Call(0x10));
        Check(vba::ReadCallsPassing(t, 16, c, 4) == 1 && c[0].argSlot == 2 && c[0].pushOp == kFLdRf,
              "pushed first, it is the callee's second slot");
        t = Place(Code().Op(kFLdAd, 8).Op(kLitI4, 1).W(kMulI4).Call(0x10));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 0, "a call with an expression argument is not read");
        t = Place(Code().Op(kFLdAd, 8).Call(0x08, 1).Op(kBos, 0).Op(kFLdAd, 8).Call(0x08, 2));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 2 && c[0].index == 1 && c[1].index == 2,
              "every call it is passed to, in body order");
        Check(vba::ReadCallsPassing(t, 16, c, 4) == 0, "a slot not passed on is in no call");
    }

    // The literal table and the names.
    Check(Is(vba::PcodeLiteralTypeName(1524), "Double") && Is(vba::PcodeLiteralTypeName(1523), "Currency") &&
          Is(vba::PcodeLiteralTypeName(1521), "Single") && Is(vba::PcodeLiteralTypeName(1653), "Integer") &&
          Is(vba::PcodeLiteralTypeName(1698), "LongLong") && Is(vba::PcodeLiteralTypeName(1999), nullptr),
          "each literal slot names the type it was measured carrying");
    Check(Is(vba::ArgTypeName("Long", true), "Long&") && Is(vba::ArgTypeName("Double&", false), "Double") &&
          Is(vba::ArgTypeName("LongLong", true), "Ref&") && Is(vba::ArgTypeName("Ref", true), "Ref&") &&
          Is(vba::ArgTypeName("Variant", false), nullptr) && Is(vba::ArgTypeName("Udt", true), nullptr),
          "names are the loads' spelling; a ByVal Variant and a record name nothing");

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
