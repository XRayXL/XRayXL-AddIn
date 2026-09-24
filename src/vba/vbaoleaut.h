#pragma once
#include <cstdint>
#include "../core/safemem.h"

// OLE Automation facts both decoders need, including the one SAFEARRAY header reader.
//
// Bounds are stored in raw rgsabound order, where rgsabound[0] is the last declared dimension;
// renderers walk them backwards.

namespace vba
{
    // ---- VARIANT ------------------------------------------------------------

    // Tag bits that change what the value is, rather than what type it is.
    constexpr std::uint16_t kVT_ARRAY = 0x2000;   // the value is a SAFEARRAY*
    constexpr std::uint16_t kVT_BYREF = 0x4000;   // the value is a pointer to one
    constexpr std::uint16_t kVT_TYPEMASK = 0x0FFF;

    // An omitted Optional arrives as a VARIANT of exactly these two, so it is safe to recognise
    // with no type information.
    constexpr std::uint16_t kVT_ERROR      = 10;
    constexpr std::uint32_t kParamNotFound = 0x80020004u;

    // nullptr for one it cannot name. VT_ARRAY and VT_BYREF are not masked off: those bits
    // change what the value is. 9 and 13 are both "Object", as VBA shows no difference.
    inline const char* VtName(std::uint16_t vt)
    {
        switch (vt)
        {
        case 0:  return "Empty";   case 1:  return "Null";
        case 2:  return "Integer"; case 3:  return "Long";
        case 4:  return "Single";  case 5:  return "Double";
        case 6:  return "Currency";case 7:  return "Date";
        case 8:  return "String";  case 9:  return "Object";
        case 10: return "Error";   case 11: return "Boolean";
        case 12: return "Variant"; case 13: return "Object";
        case 14: return "Decimal"; case 17: return "Byte";
        case 20: return "LongLong";
        // Only a COM property produces these, so named by width rather than an invented VBA
        // name. 36 is a user-defined Type carried in a Variant.
        case 16: return "Int8";    case 18: return "UInt16";
        case 19: return "UInt32";  case 21: return "UInt64";
        case 22: return "Int";     case 23: return "UInt";
        case 36: return "Udt";
        default: return nullptr;
        }
    }

    // VtName with "()": static strings a caller may keep.
    inline const char* ArrayTypeName(std::uint16_t vt)
    {
        switch (vt)
        {
        case 0:  return "Empty()";   case 1:  return "Null()";
        case 2:  return "Integer()"; case 3:  return "Long()";
        case 4:  return "Single()";  case 5:  return "Double()";
        case 6:  return "Currency()";case 7:  return "Date()";
        case 8:  return "String()";  case 9:  return "Object()";
        case 10: return "Error()";   case 11: return "Boolean()";
        case 12: return "Variant()"; case 13: return "Object()";
        case 14: return "Decimal()"; case 17: return "Byte()";
        case 20: return "LongLong()";
        case 16: return "Int8()";    case 18: return "UInt16()";
        case 19: return "UInt32()";  case 21: return "UInt64()";
        case 22: return "Int()";     case 23: return "UInt()";
        case 36: return "Udt()";
        default: return "?()";
        }
    }

    // ---- SAFEARRAY ----------------------------------------------------------

    // Every documented FADF_ bit; anything else set means it is not a SAFEARRAY.
    constexpr std::uint16_t kFadfKnown    = 0x0FF7;
    constexpr std::uint16_t kFadf_Record  = 0x0020, kFadf_HaveVt   = 0x0080,
                            kFadf_BSTR    = 0x0100, kFadf_Unknown  = 0x0200,
                            kFadf_Dispatch= 0x0400, kFadf_Variant  = 0x0800;

    struct SaInfo
    {
        std::uint16_t cDims = 0, fFeat = 0, vt = 0;
        std::uint32_t cbElem = 0;
        std::uint64_t pvData = 0;
        std::uint32_t cElems[8] = {};
        std::int32_t  lBound[8] = {};
        std::uint64_t total = 0;
    };

    // A feature bit is more specific than the VARTYPE, so it overrides it. 0: nothing named one.
    inline std::uint16_t EffectiveElemVt(const SaInfo& s)
    {
        std::uint16_t vt = s.vt & kVT_TYPEMASK;
        // An Enum array says VT_USERDEFINED and stores Longs; TypeName says Long() too.
        if (vt == 29 && s.cbElem == 4) vt = 3;
        if (s.fFeat & kFadf_Unknown)  vt = 13;   // IUnknown*
        if (s.fFeat & kFadf_Dispatch) vt = 9;    // IDispatch*
        if (s.fFeat & kFadf_BSTR)     vt = 8;
        if (s.fFeat & kFadf_Variant)  vt = 12;   // 24-byte self-typed elements
        return vt;
    }

    // `vtHint` is the element type a caller already knew from the VARIANT that held this array,
    // used only when the descriptor carries none; pass 0 otherwise. False for anything that
    // does not prove itself an array.
    inline bool ReadSafeArrayHeader(std::uint64_t psa, std::uint16_t vtHint, SaInfo& s)
    {
        s = SaInfo{};
        if (!core::InRangeAndAligned(psa, 8)) return false;

        std::uint16_t cDims = 0, fFeat = 0;
        std::uint32_t cbElem = 0, cLocks = 0;
        std::uint64_t pv = 0;
        if (!core::RdU16(psa + 0x00, cDims))  return false;
        if (!core::RdU16(psa + 0x02, fFeat))  return false;
        if (!core::RdU32(psa + 0x04, cbElem)) return false;
        if (!core::RdU32(psa + 0x08, cLocks)) return false;
        if (!core::RdU64(psa + 0x10, pv))     return false;

        if (cDims == 0 || cDims > 8)              return false;
        // No feature bits is legal when the holder already names the element type.
        if ((fFeat & ~kFadfKnown) || (fFeat == 0 && vtHint == 0)) return false;
        // Range-checked, not held to scalar widths: a record array's element is sizeof(the UDT).
        if (cbElem == 0 || cbElem > 0x10000)      return false;
        // A large lock count means a stale structure; a sanity bound, not a documented limit.
        if (cLocks > 0x1000)                      return false;

        s.cDims = cDims; s.fFeat = fFeat; s.cbElem = cbElem; s.pvData = pv;
        s.total = 1;
        for (std::uint16_t d = 0; d < cDims; ++d)
        {
            std::uint32_t n = 0, lb = 0;
            if (!core::RdU32(psa + 0x18 + static_cast<std::uint64_t>(d) * 8,     n))  return false;
            if (!core::RdU32(psa + 0x18 + static_cast<std::uint64_t>(d) * 8 + 4, lb)) return false;
            // A zero count is valid: `Array()` makes one, with bounds 0..-1.
            if (n > 0x4000000) return false;                       // 64M per dimension
            // Or a stale lLbound makes "lower + count - 1" overflow.
            const std::int32_t low = static_cast<std::int32_t>(lb);
            if (low < -0x10000000 || low > 0x10000000) return false;
            s.cElems[d] = n;
            s.lBound[d] = low;
            s.total *= n;
            if (s.total > 0x8000000ull) return false;               // 128M total
        }

        // An empty array may have a null data pointer.
        if (s.total > 0 && !core::InUserRange(pv)) return false;

        // With FADF_HAVEVARTYPE the vartype is the DWORD before the descriptor, as
        // SafeArrayGetVartype reads it; the hint stands when that is missing or unreadable.
        s.vt = vtHint;
        if (fFeat & kFadf_HaveVt)
        {
            std::uint32_t vt = 0;
            if (core::RdU32(psa - 4, vt) && (vt & kVT_TYPEMASK) != 0 && vt <= 0xFFFF)
                s.vt = static_cast<std::uint16_t>(vt);
        }

        // Everything above is only plausibility; a VBA array always names its element type.
        const std::uint16_t evt = EffectiveElemVt(s);
        if (evt == 0) return false;

        // The one real cross-check, and why the loose width range above is safe.
        std::uint32_t want = 0;
        switch (evt & kVT_TYPEMASK)
        {
        case 2: case 11:                     want = 2;  break;   // I2, BOOL
        case 3: case 4: case 10:             want = 4;  break;   // I4, R4, ERROR
        case 5: case 6: case 7: case 20:     want = 8;  break;   // R8, CY, DATE, I8
        case 8: case 9: case 13:             want = 8;  break;   // BSTR, IDispatch*, IUnknown*
        case 12:                             want = 24; break;   // VARIANT on x64
        case 14:                             want = 16; break;   // DECIMAL
        case 16: case 17:                    want = 1;  break;   // I1, UI1
        case 18:                             want = 2;  break;   // UI2
        case 19: case 22: case 23:           want = 4;  break;   // UI4, INT, UINT
        case 21:                             want = 8;  break;   // UI8
        default: break;                                          // RECORD and the rest: no fixed width
        }
        if (want && want != cbElem) return false;
        return true;
    }
}
