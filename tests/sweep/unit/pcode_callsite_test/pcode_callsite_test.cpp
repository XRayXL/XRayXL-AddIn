// Unit test for reading a caller's call site (vba/vbapcode.cpp): the pushes straight before a
// call, each one operand-stack slot, read back from the call to the first that is not one; the
// calls a parameter is passed on to; and the type each literal slot carries.
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
    constexpr std::uint16_t kFLdI8 = 667;

    struct Code
    {
        std::vector<std::uint8_t> b;
        Code& W(std::uint16_t v) { b.push_back(static_cast<std::uint8_t>(v)); b.push_back(static_cast<std::uint8_t>(v >> 8)); return *this; }
        Code& D(std::int32_t v) { for (int i = 0; i < 4; ++i) b.push_back(static_cast<std::uint8_t>(v >> (8 * i))); return *this; }
        Code& Op(std::uint16_t op, std::int32_t operand) { return W(op).D(operand); }
        Code& Str() { return W(kLitStr).W(0); }
        Code& Call(std::uint16_t argBytes, std::uint16_t index = 0) { return W(kCall).W(index).W(argBytes); }
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
    for (std::uint16_t op : { kBos, kLitI4, kFLdRf, kFLdAd, kFLdI4, kFLdFPR8, kCall, kFLdI8 }) L.len[op] = 6;
    L.len[1653] = 2;
    L.len[kLitStr] = 4; L.len[kLitR8] = 10; L.len[kMulI4] = 2;
    vba::SetArmedLengths(L);
    vba::PcodePush p[4] = {};

    // `Callee 42, "hi", n`: pushed last to first, so the local's address is pushed first.
    {
        const std::uint64_t t = Place(Code().Op(kFLdRf, -16).Str().Op(kLitI4, 42).Call(0x18));
        const bool ok = vba::ReadCallPushes(t, g_callOff, 3, p) == 3;
        Check(ok && p[0].kind == vba::PushKind::SlotAddress && p[0].operand == -16 &&
              p[1].kind == vba::PushKind::Literal && Is(p[1].type, "String") &&
              p[2].kind == vba::PushKind::Literal && Is(p[2].type, "Long"),
              "three pushes read in order: an address, a string, a long");
    }
    {
        const std::uint64_t t = Place(Code().Op(kFLdI4, -8).Op(kFLdAd, 8).Call(0x10));
        Check(vba::ReadCallPushes(t, g_callOff, 2, p) == 2 && p[0].kind == vba::PushKind::Value &&
              Is(p[0].type, "Long") && p[1].kind == vba::PushKind::HeldPointer && p[1].operand == 8,
              "a typed load is a value, a ByRef parameter's pointer is held");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 1).Op(kLitI4, 2).Call(0x08));
        Check(vba::ReadCallPushes(t, g_callOff, 1, p) == 1 && p[0].operand == 2,
              "only the last n instructions are the call's: `x = a + F(b)` pushes a first");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 12).Op(kFLdI4, -16).W(kMulI4).Call(0x08));
        Check(vba::ReadCallPushes(t, g_callOff, 1, p) == 0, "an expression in slot 1 reads nothing");
    }
    {
        const std::uint64_t t = Place(Code().Op(kFLdFPR8, -8).Call(0x08));
        Check(vba::ReadCallPushes(t, g_callOff, 1, p) == 0, "an FPU load is not a push");
    }
    {
        const std::uint64_t t = Place(Code().Op(kLitI4, 3).Call(0x10));
        Check(vba::ReadCallPushes(t, g_callOff, 2, p) == -1, "fewer instructions than slots is a misread");
    }

    // Expressions: slots from 1 are read back to the first instruction that is not a push.
    {
        // `F a, b * 2, c` with a ByRef local a: c first, then b * 2, then a
        const std::uint64_t t = Place(Code().Op(kFLdI4, -24).Op(kFLdI4, -16).Op(kLitI4, 2).W(kMulI4)
                                            .Op(kFLdRf, -8).Call(0x18));
        const int r = vba::ReadCallPushes(t, g_callOff, 3, p);
        Check(r == 1 && p[2].kind == vba::PushKind::SlotAddress && p[2].operand == -8,
              "`F a, b * 2, c`: slot 1 read, the expression in slot 2 ends it, slot 3 unread");
    }
    {
        // `F a, b, c * 2`: the expression is pushed first, so slots 1 and 2 are read
        const std::uint64_t t = Place(Code().Op(kFLdI4, -24).Op(kLitI4, 2).W(kMulI4)
                                            .Op(kFLdI4, -16).Op(kFLdAd, 8).Call(0x18));
        const int r = vba::ReadCallPushes(t, g_callOff, 3, p);
        Check(r == 2 && p[2].kind == vba::PushKind::HeldPointer && p[1].kind == vba::PushKind::Value &&
              p[1].operand == -16, "`F a, b, c * 2`: slots 1 and 2 read");
    }
    {
        // `F G(x), y`: y is pushed, then G's argument and G; F's slot 1 is G's result
        Code c;
        c.Op(kFLdI4, -16).Op(kFLdRf, -8).Call(0x08, 7).W(kCall).W(8).W(0x10);   // then F, pool 8, two slots
        const std::uint64_t t = Place(c);
        const std::uint32_t fAt = g_callOff + 6;            // straight after G's call
        Check(vba::ReadCallPushes(t, g_callOff, 1, p) == 1 && p[0].operand == -8,
              "`F G(x), y`: G's own slot 1 is x");
        Check(vba::ReadCallPushes(t, fAt, 2, p) == 0, "`F G(x), y`: F's slot 1 is G's result, not a push");
    }
    {
        // `F G(x), y` again, now as `F y, G(x)`: G runs first, then y is pushed last
        Code c;
        c.Op(kFLdRf, -8).Call(0x08, 7).Op(kFLdI4, -16).W(kCall).W(8).W(0x10);
        const std::uint64_t t = Place(c);
        const std::uint32_t fAt = g_callOff + 12;           // past G's call and y's push
        const int r = vba::ReadCallPushes(t, fAt, 2, p);
        Check(r == 1 && p[1].operand == -16, "`F y, G(x)`: slot 1 is y, slot 2 is G's result and unread");
    }

    // What a push cannot say: a local's eight bytes may be a temporary's pointer. Still one push.
    {
        const std::uint64_t t = Place(Code().Op(kFLdI8, 16).Op(kFLdI8, -208).Call(0x10));
        Check(vba::ReadCallPushes(t, g_callOff, 2, p) == 2 && Is(p[1].type, nullptr) &&
              Is(p[0].type, "LongLong"),
              "eight bytes from a local name nothing, from a ByVal parameter its LongLong");
    }

    // A store that sets a ByRef Variant as readily as a ByRef named class names neither.
    Check(Is(vba::PcodeStoreTypeName(782), nullptr) && Is(vba::PcodeTypeName(750), "Object&") &&
          Is(vba::PcodeStoreTypeName(783), "Variant&"), "782 names nothing; 750 still reads Object&");

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
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 0, "`G x * 1, p`: an expression between it and the call hides it");
        t = Place(Code().Op(kLitI4, 3).Op(kLitI4, 1).W(kMulI4).Op(kFLdAd, 8).Call(0x10));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 1 && c[0].argSlot == 1,
              "`G p, 3 * 1`: an expression pushed before it does not");
        t = Place(Code().Op(kLitI4, 3).Op(kLitI4, 1).W(kMulI4).Op(kFLdAd, 8).Op(kLitI4, 4).Call(0x18));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 1 && c[0].argSlot == 2,
              "`G 4, p, 3 * 1`: it lands in slot 2");
        t = Place(Code().Op(kFLdAd, 8).Op(kFLdAd, 8).Call(0x10));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 1 && c[0].argSlot == 1, "`G p, p`: once, at slot 1");
        t = Place(Code().Op(kFLdAd, 8).Call(0x10));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 0, "two slots and one instruction: a misread passes nothing");
        t = Place(Code().Op(kFLdAd, 8).Call(0x08, 1).Op(kBos, 0).Op(kFLdAd, 8).Call(0x08, 2));
        Check(vba::ReadCallsPassing(t, 8, c, 4) == 2 && c[0].index == 1 && c[1].index == 2,
              "every call it is passed to, in body order");
        Check(vba::ReadCallsPassing(t, 16, c, 4) == 0, "a slot not passed on is in no call");
    }

    // The literal table and the names.
    Check(Is(vba::PcodeLiteralTypeName(1524), "Double") && Is(vba::PcodeLiteralTypeName(1523), "Currency") &&
          Is(vba::PcodeLiteralTypeName(1521), "Single") && Is(vba::PcodeLiteralTypeName(1698), "LongLong") &&
          Is(vba::PcodeLiteralTypeName(1999), nullptr) && !vba::PcodeIsLiteral(1999),
          "each literal slot names the type it was measured carrying");
    Check(Is(vba::PcodeLiteralTypeName(1653), nullptr) && Is(vba::PcodeLiteralTypeName(1517), nullptr) &&
          vba::PcodeIsLiteral(1653) && vba::PcodeIsLiteral(1517),
          "an Integer literal is a literal of no type: it also fills a Byte parameter");
    {
        // `F 5, n` into (ByVal b As Byte, ByVal n As Long): read past the literal, typed after it
        const std::uint64_t t = Place(Code().Op(kFLdI4, 8).W(1653).Call(0x10));
        Check(vba::ReadCallPushes(t, g_callOff, 2, p) == 2 && Is(p[1].type, nullptr) &&
              p[1].kind == vba::PushKind::Literal && Is(p[0].type, "Long"),
              "an Integer literal is one push: the slot after it is still read");
    }
    Check(Is(vba::ArgTypeName("Long", true), "Long&") && Is(vba::ArgTypeName("Double&", false), "Double") &&
          Is(vba::ArgTypeName("LongLong", true), "Ref&") && Is(vba::ArgTypeName("Ref", true), "Ref&") &&
          Is(vba::ArgTypeName("Variant", false), nullptr) && Is(vba::ArgTypeName("Udt", true), nullptr),
          "names are the loads' spelling; a ByVal Variant and a record name nothing");

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
