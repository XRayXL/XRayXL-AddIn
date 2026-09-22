#include "vbaretdecode.h"
#include "vbaoleaut.h"
#include "core/excelerr.h"
#include "core/render.h"
#include "vbaobject.h"

#include <windows.h>
#include <cstdio>
#include <cstring>
#include "core/safemem.h"

namespace vba
{
    namespace
    {
        // Readers and address tests: core/safemem.h. Every byte read here is
        // VBE7's, and a stale pointer must cost the read and nothing more. A BSTR is
        // 2-aligned: it points four bytes past its length prefix, so demanding 8
        // would refuse real strings.
        using core::InUserRange;
        using core::InRangeAndAligned;
        using core::RdU64;
        using core::RdU32;
        using core::RdU16;
        using core::RdU8;
        using core::RdWide;
        using core::ValueWriter;

        // The only BSTR reader, and strict: a wrong guess prints an unrelated allocation as a
        // value. Truncation past kBstrMaxChars is marked, beyond kBstrMaxBytes the whole thing
        // is refused, and a control character other than tab means the memory is not text.
        constexpr std::uint32_t kBstrMaxBytes = 4096;   // refuse beyond this when PROBING
        constexpr int           kBstrMaxChars = 256;    // then show this many

        // When the type is already known to be BSTR, long strings are allowed.
        constexpr std::uint32_t kBstrMaxBytesTold = 1u << 20;   // 512K characters

        // `told` -- does something ELSE already say this is a string? It decides
        // exactly one thing: whether a ZERO length is a value. When a VARIANT tag or
        // an exit opcode says BSTR, `cb == 0` is an empty string. When PROBING an
        // untyped slot, `cb == 0` is also what zeroed memory looks like, and is
        // refused.
        bool DescribeBstr(std::uint64_t p, ValueWriter& w, bool told)
        {
            if (!InRangeAndAligned(p, 2)) return false;
            std::uint32_t cb = 0;
            if (!RdU32(p - 4, cb)) return false;
            if (cb & 1) return false;
            if (cb > (told ? kBstrMaxBytesTold : kBstrMaxBytes)) return false;
            const int n = static_cast<int>(cb / 2);
            if (n == 0)
            {
                if (!told) return false;          // zeroed memory looks like this
                w.String(L"", 0, false);
                return true;
            }

            wchar_t text[kBstrMaxChars + 2] = {};
            const int take = n < kBstrMaxChars ? n : kBstrMaxChars;
            if (!RdWide(p, text, take)) return false;

            // NUL-terminated PAST its full length -- checked at the real end,
            // which is why the byte limit and the character cap are separate.
            std::uint16_t term = 1;
            if (!RdU16(p + cb, term) || term != 0) return false;

            // A control character is a refusal only when probing. When a VARIANT tag, an exit
            // opcode or a declared `As String` has already said BSTR, the same byte is a
            // `vbCrLf` in a message or a tab in a record.
            if (!told)
                for (int i = 0; i < take; ++i)
                    if (text[i] < 32 && text[i] != 9) return false;

            w.String(text, take, n > take);
            return true;
        }

        bool DescribeSafeArray(std::uint64_t psa, std::uint16_t vtHint, ValueWriter& w,
                               int depth, std::uint16_t* vtOut = nullptr);
        bool DescribeVariantAt(std::uint64_t at, ValueWriter& w, int depth);

        // Nesting is bounded: each level is a real recursion on the hot-path stack and the
        // memory could be cyclic. Deeper than this reads as "[...]". The counter advances on
        // the Variant-in-a-Variant hop, in DescribeElement's VT_VARIANT case, not on the array.
        constexpr int kMaxNest = 32;

        // The one place an object is rendered, so an argument and a result read the same. With
        // OBJECTS on the class is named and a Range, Worksheet or Workbook is described;
        // anything that does not work out falls back to the address.
        volatile LONG g_describeObjects = 0;

        void DescribeObject(std::uint64_t ptr, ValueWriter& w)
        {
            if (ptr == 0) { w.Word("Nothing"); return; }

            // The address is always there: it is how one object is followed from an argument to
            // a result, which a class name cannot do.
            if (InterlockedCompareExchange(&g_describeObjects, 0, 0) != 0 &&
                DescribeObjectDetail(ptr, w))
                return;
            w.BeginObject(nullptr, ptr, nullptr);
            w.EndObject();
        }

        // A DECIMAL: a 96-bit magnitude, a scale of 0..28 and a sign byte, laid over the whole
        // VARIANT: scale at +2, sign at +3, the high 32 bits at +4, the low 64 at +8. Rendered
        // exactly, in integer arithmetic, since 28 digits is more than a double holds. A scale
        // above 28 or a sign byte other than 0 or 0x80 is refused.
        bool DescribeDecimal(std::uint64_t at, ValueWriter& w)
        {
            char out[40];
            constexpr int cap = static_cast<int>(sizeof(out));
            std::uint8_t scale = 0, sign = 0; std::uint32_t hi = 0; std::uint64_t lo = 0;
            if (!RdU8(at + 2, scale) || !RdU8(at + 3, sign) ||
                !RdU32(at + 4, hi) || !RdU64(at + 8, lo)) return false;
            if (scale > 28 || (sign != 0 && sign != 0x80)) return false;

            // Digits least-significant first, by dividing the 96-bit value by
            // ten until it is gone; then at least one digit left of the point.
            std::uint32_t limb[3] = { static_cast<std::uint32_t>(lo),
                                      static_cast<std::uint32_t>(lo >> 32), hi };
            char digits[32]; int n = 0;
            while (n < 30 && (limb[0] | limb[1] | limb[2]))
            {
                std::uint64_t rem = 0;
                for (int i = 2; i >= 0; --i)
                {
                    const std::uint64_t cur = (rem << 32) | limb[i];
                    limb[i] = static_cast<std::uint32_t>(cur / 10);
                    rem = cur % 10;
                }
                digits[n++] = static_cast<char>('0' + rem);
            }
            if (n == 0) digits[n++] = '0';
            while (n <= scale) digits[n++] = '0';

            int j = 0;
            if (sign && j < cap - 1) out[j++] = '-';
            for (int i = n - 1; i >= 0 && j < cap - 1; --i)
            {
                out[j++] = digits[i];
                if (i == scale && scale > 0 && j < cap - 1) out[j++] = '.';
            }
            out[j] = 0;
            w.Number("Decimal", out);
            return true;
        }

        // A number of VBA type `type`, formatted from `fmt`.
        template <class T>
        void PutNumber(ValueWriter& w, const char* type, const char* fmt, T v)
        {
            char t[48];
            _snprintf_s(t, _TRUNCATE, fmt, v);
            w.Number(type, t);
        }

        // One element of the given VARTYPE, read from `at`. `depth` is how
        // far inside nested arrays this element sits. False, having written
        // nothing, when the bytes cannot be read.
        bool DescribeElement(std::uint16_t vt, std::uint64_t at, ValueWriter& w, int depth = 0)
        {
            std::uint64_t q = 0; std::uint32_t d = 0; std::uint16_t h = 0;
            switch (vt)
            {
            case 2:  // VT_I2 -- Integer AND Boolean share it
                if (!RdU16(at, h)) return false;
                PutNumber(w, "Integer", "%d", static_cast<int>(static_cast<std::int16_t>(h)));
                return true;
            case 3:  // VT_I4
                if (!RdU32(at, d)) return false;
                PutNumber(w, "Long", "%d", static_cast<int>(static_cast<std::int32_t>(d)));
                return true;
            case 4:  // VT_R4
            {
                if (!RdU32(at, d)) return false;
                float f = 0; memcpy(&f, &d, sizeof(f));
                PutNumber(w, "Single", "%.9g", static_cast<double>(f));
                return true;
            }
            case 5:  // VT_R8
            {
                if (!RdU64(at, q)) return false;
                double v = 0; memcpy(&v, &q, sizeof(v));
                w.Double(v);
                return true;
            }
            case 7:  // VT_DATE -- a serial number, which is what a cell shows
            {
                if (!RdU64(at, q)) return false;
                double v = 0; memcpy(&v, &q, sizeof(v));
                char t[48];
                core::FormatDouble(v, t, sizeof t);
                w.Number("Date", t);
                return true;
            }
            case 6:  // VT_CY -- a 64-bit integer scaled by 10,000
            {
                // EXACT, IN INTEGER ARITHMETIC, with the sign carried explicitly. Currency
                // has 19 significant digits, which a double cannot hold, and dividing first
                // truncates toward zero, which loses the sign of anything in (-1, 0). A
                // wrong-signed money value is the worst decoding error this file can make.
                if (!RdU64(at, q)) return false;
                const std::int64_t c = static_cast<std::int64_t>(q);
                const bool neg = (c < 0);
                const std::uint64_t a =
                    neg ? (~static_cast<std::uint64_t>(c) + 1u)
                        : static_cast<std::uint64_t>(c);
                char t[48];
                _snprintf_s(t, _TRUNCATE, "%s%llu.%04llu", neg ? "-" : "",
                            static_cast<unsigned long long>(a / 10000),
                            static_cast<unsigned long long>(a % 10000));
                w.Number("Currency", t);
                return true;
            }
            case 8:  // VT_BSTR
                if (!RdU64(at, q)) return false;
                if (q == 0) { w.String(L"", 0, false); return true; }
                return DescribeBstr(q, w, /*told=*/true);
            case 11: // VT_BOOL: VARIANT_TRUE is -1. Spelt as Excel spells it, like an error value.
                if (!RdU16(at, h)) return false;
                w.Bool(h != 0);
                return true;
            case 17: // VT_UI1 -- one byte, read as one byte (see RdU8)
            {
                std::uint8_t b = 0;
                if (!RdU8(at, b)) return false;
                PutNumber(w, "Byte", "%u", static_cast<unsigned>(b));
                return true;
            }
            case 20: // VT_I8
                if (!RdU64(at, q)) return false;
                PutNumber(w, "LongLong", "%lld", static_cast<long long>(static_cast<std::int64_t>(q)));
                return true;
            case 14: // VT_DECIMAL: in an array the element IS the whole DECIMAL
                return DescribeDecimal(at, w);
            // THE FOREIGN INTEGER WIDTHS. No VBA declaration produces them, but a
            // COM property can hand one back inside a Variant, and a width is a
            // width: each is read at its size and printed as the number it is.
            case 16: // VT_I1
            {
                std::uint8_t b = 0;
                if (!RdU8(at, b)) return false;
                PutNumber(w, VtName(vt), "%d", static_cast<int>(static_cast<std::int8_t>(b)));
                return true;
            }
            case 18: // VT_UI2
                if (!RdU16(at, h)) return false;
                PutNumber(w, VtName(vt), "%u", static_cast<unsigned>(h));
                return true;
            case 19: case 23: // VT_UI4, VT_UINT
                if (!RdU32(at, d)) return false;
                PutNumber(w, VtName(vt), "%u", static_cast<unsigned>(d));
                return true;
            case 22: // VT_INT
                if (!RdU32(at, d)) return false;
                PutNumber(w, VtName(vt), "%d", static_cast<int>(static_cast<std::int32_t>(d)));
                return true;
            case 21: // VT_UI8
                if (!RdU64(at, q)) return false;
                PutNumber(w, VtName(vt), "%llu", static_cast<unsigned long long>(q));
                return true;
            case 9: case 13:  // VT_DISPATCH, VT_UNKNOWN
                if (!RdU64(at, q)) return false;
                DescribeObject(q, w);
                return true;
            case 10: // VT_ERROR: a cell error travelling in an array (Range.Value)
            {
                if (!RdU32(at, d)) return false;
                // SPELT AS EXCEL SPELLS IT. `#N/A` is what the user sees in the
                // cell; `Error(0x7FA)` made them do the arithmetic. Unquoted, so
                // it cannot be confused with the string "#N/A", which renders
                // with its quotes.
                if (const char* en = core::ExcelErrNameFromVba(d)) w.Error(en);
                else
                {
                    char t[32];
                    _snprintf_s(t, _TRUNCATE, "Error(0x%X)", d);
                    w.Error(t);
                }
                return true;
            }
            case 12: // VT_VARIANT: the element IS a VARIANT, self-typed -- recurse
                return DescribeVariantAt(at, w, depth + 1);
            default:
            {
                // An element type this cannot read is named, not invented: a
                // placeholder says "there was one here, of type N" and leaves
                // the rest of the array intact.
                char t[24];
                _snprintf_s(t, _TRUNCATE, "?vt%u", static_cast<unsigned>(vt));
                w.Marker(t);
                return true;
            }
            }
        }


        // The one array renderer, for both columns: the element type, both bounds of every
        // dimension in declaration order, then one brace level per dimension with the last
        // index varying fastest. The bytes lie column-major and rgsabound is stored right to
        // left, so each element's position is computed from its indices.
        //
        // `sa` must already have come from ReadSafeArrayHeader, which makes the bounds and the
        // element count safe to walk.
        bool RenderSaText(const SaInfo& sa, std::uint16_t vt, ValueWriter& w, int depth)
        {
            const int dims = sa.cDims > 8 ? 8 : static_cast<int>(sa.cDims);
            long long lo[8] = {}, hi[8] = {};
            std::uint64_t extent[8] = {}, stride[8] = {};
            for (int d = 0; d < dims; ++d)
            {
                const int raw = dims - 1 - d;                     // rgsabound is reversed
                extent[d] = sa.cElems[raw];
                lo[d] = static_cast<long long>(sa.lBound[raw]);
                hi[d] = lo[d] + static_cast<long long>(extent[d]) - 1;
                stride[d] = d ? stride[d - 1] * extent[d - 1] : 1;
            }
            w.BeginArray(VtName(vt), dims, lo, hi);
            core::WalkRowMajor(w, dims, extent, [&](const std::uint64_t* idx)
            {
                std::uint64_t at = 0;
                for (int d = 0; d < dims; ++d) at += idx[d] * stride[d];
                // One element the decoder will not vouch for is `?`, not a refusal of the whole
                // array: the shape came from the descriptor and stays true.
                const ValueWriter::Mark m = w.Save();
                if (!DescribeElement(vt, sa.pvData + at * sa.cbElem, w, depth))
                {
                    w.Restore(m);
                    w.Marker("?");
                }
            });
            w.EndArray();
            return true;
        }


        bool DescribeSafeArray(std::uint64_t psa, std::uint16_t vtHint, ValueWriter& w,
                               int depth, std::uint16_t* vtOut)
        {
            if (vtOut) *vtOut = 0;
            if (depth > kMaxNest) { w.Marker("[...]"); return true; }

            // One SAFEARRAY reader for both columns (vbaoleaut.h).
            SaInfo sa{};
            if (!ReadSafeArrayHeader(psa, vtHint, sa)) return false;
            const std::uint16_t vt = EffectiveElemVt(sa);
            // REPORTED BACK, so a caller that needs to NAME the element type does not
            // read the descriptor a second time.
            if (vtOut) *vtOut = vt;
            return RenderSaText(sa, vt, w, depth);
        }

        // Is this a VARTYPE the decoder can read a value for? Consulted only at the top level
        // of a result: inside an array an unreadable element renders `?vtN` and the array
        // survives, but as the whole result it would claim a Variant was returned from bytes
        // whose own tag says otherwise. The row goes empty and is counted.
        bool ReadableVartype(std::uint16_t vt)
        {
            switch (vt)
            {
            case 0: case 1:                       // Empty, Null
            case 2: case 3: case 4: case 5:       // I2, I4, R4, R8
            case 6: case 7: case 8:               // CY, DATE, BSTR
            case 9: case 10: case 11:             // DISPATCH, ERROR, BOOL
            case 12: case 13: case 14:            // VARIANT, UNKNOWN, DECIMAL
            case 16: case 17: case 18: case 19:   // I1, UI1, UI2, UI4
            case 20: case 21: case 22: case 23:   // I8, UI8, INT, UINT
            case 36:                              // RECORD: a UDT, as its address
                return true;
            default:
                return false;
            }
        }

        // The held value of the VARIANT whose tag is `vt` and whose payload is at `val`.
        bool VariantBody(std::uint16_t vt, std::uint64_t val, std::uint64_t self, bool byref,
                         ValueWriter& w, int depth)
        {
            if (vt & kVT_ARRAY)                   // the element type is in the low bits
            {
                std::uint64_t psa = 0;
                if (!RdU64(val, psa)) return false;
                // A null SAFEARRAY is a dynamic array never allocated: `Dim a() As Long`.
                if (psa == 0) { w.Unallocated(VtName(static_cast<std::uint16_t>(vt & kVT_TYPEMASK))); return true; }
                return DescribeSafeArray(psa, static_cast<std::uint16_t>(vt & kVT_TYPEMASK), w, depth);
            }
            switch (vt)
            {
            case 0:  w.Word("Empty"); return true;
            case 1:  w.Word("Null");  return true;
            case 10: // VT_ERROR: the SCODE, which for a UDF is a cell error
                return DescribeElement(10, val, w, depth);
            case 9: case 13:                       // objects
            {
                std::uint64_t ptr = 0;
                if (!RdU64(val, ptr)) return false;
                DescribeObject(ptr, w);
                return true;
            }
            case 14:                               // the DECIMAL overlays the VARIANT itself
                return DescribeDecimal(self, w);
            case 36:                               // VT_RECORD: a user-defined Type
            {
                // Its layout lives behind IRecordInfo, a COM call this hook does
                // not make, so it is its address -- the same shape the argument
                // column gives a UDT, so one can be followed between rows.
                //
                // After a by-reference hop, val already is the record's address.
                std::uint64_t rec = val;
                if (!byref && !RdU64(val, rec)) return false;
                w.Udt(rec);
                return true;
            }
            default:
                return DescribeElement(vt, val, w, depth);
            }
        }

        // A VARIANT at `at`: vt, then the value at +8, decoded by vt. Written inside
        // BeginVariant/EndVariant, and nothing at all when it fails.
        bool DescribeVariantAt(std::uint64_t at, ValueWriter& w, int depth)
        {
            std::uint16_t vt = 0;
            if (!RdU16(at, vt)) return false;
            std::uint64_t val  = at + 8;          // the payload
            std::uint64_t self = at;              // where a DECIMAL overlays the VARIANT
            // True after a by-reference hop; VT_RECORD needs to know.
            bool byref = false;

            // VT_BYREF: the VARIANT holds a pointer to the value, which is how the interpreter
            // passes a typed variable into a Variant parameter (0x4008 is VT_BYREF|VT_BSTR).
            // Followed once, under the same guard as every other pointer, and at the top level
            // only.
            if (vt & kVT_BYREF)
            {
                std::uint64_t target = 0;
                if (depth != 0 || !RdU64(val, target) || !InUserRange(target)) return false;
                vt   = static_cast<std::uint16_t>(vt & ~kVT_BYREF);
                val  = self = target;
                byref = true;
                if (vt == 12)                     // VT_BYREF|VT_VARIANT: a VARIANT elsewhere
                {
                    if (!RdU16(target, vt) || (vt & kVT_BYREF)) return false;   // one hop
                    val = target + 8;
                }
            }
            if (!(vt & kVT_ARRAY) && depth == 0 && !ReadableVartype(vt)) return false;

            const ValueWriter::Mark m = w.Save();
            w.BeginVariant();
            if (!VariantBody(vt, val, self, byref, w, depth)) { w.Restore(m); return false; }
            w.EndVariant();
            return true;
        }
    }

    // The type from the store, when the exit opcode does not carry one. Class and form
    // Functions all leave through exit opcode 1664 whatever they return, so the instruction
    // that wrote the result carries the type. The store family is the load family plus 32;
    // String, Object and Variant do not follow that pairing. Boolean needs no entry: VBA stores
    // it as an Integer.
    //
    // 694 is here, but the caller identifies a Variant from the operand (-0x18), because a
    // Variant is stored through a different opcode depending on what it holds (707 for a
    // string). See DecideReturnKind in vbatrace.cpp.
    RetKind StoreReturnKind(std::uint16_t storeOp)
    {
        switch (storeOp)
        {
        case 688: return RetKind::Byte;
        case 689: return RetKind::Integer;
        case 690: return RetKind::Long;
        case 693: return RetKind::Currency;
        case 699: return RetKind::LongLongOrArray;
        // 671 is the typed-array store, here for the same reason 699 is: on the class and form
        // path the exit opcode is 1664, which carries no type. Known gap: a class Function
        // returning a typed array still reports an empty ret.
        case 671: return RetKind::LongLongOrArray;
        case 700: return RetKind::Single;
        case 701: return RetKind::Double;

        // String, Variant and Object. Each stores where the decoder already reads it:
        //
        //    708 stores at operand -8     -- where String reads its BSTR
        //    694 stores at operand -0x18  -- where Variant reads, and nowhere else does
        //    696 stores at operand -8     -- and 696-32 = 664, the load table's Object
        //
        // A Variant holding a string stores through 707 at the same -0x18, which is why the
        // caller decides a Variant by its operand and not by this table.
        case 708: return RetKind::String;
        case 694: return RetKind::Variant;
        case 696: return RetKind::Object;
        // The same slots written from a temporary: 695 moves a computed string in (`F = a & b`,
        // `F = Trim(x)`), 706 sets an object from a function. Across the corpus each writes
        // [R14-8] only in Functions of that type.
        case 695: return RetKind::String;
        case 706: return RetKind::Object;

        default:  return RetKind::Unknown;
        }
    }


    // LATCHED AT ARM, like every other setting: the hot path never sees one
    // change under it.
    void SetDescribeObjects(bool on)
    {
        InterlockedExchange(&g_describeObjects, on ? 1 : 0);
    }

    // See vbaretdecode.h. The whole array, for a caller that read the header
    // itself -- the arguments column, which follows an extra indirection to find
    // the descriptor and so cannot use DescribeSafeArray's entry point.
    bool RenderSafeArrayValue(const SaInfo& sa, ValueWriter& w)
    {
        return RenderSaText(sa, EffectiveElemVt(sa), w, 0);
    }

    // See vbaretdecode.h. One element of a SAFEARRAY, for a caller that walked
    // the descriptor itself.
    bool DescribeArrayElement(std::uint16_t vt, std::uint64_t at, ValueWriter& w)
    {
        if (vt == 0) return false;
        return DescribeElement(vt, at, w, 0);
    }

    // See vbaretdecode.h. The ARGS column's decoder, which is this one.
    bool DescribeVariantValue(std::uint64_t at, ValueWriter& w)
    {
        return DescribeVariantAt(at, w, 0);
    }

    // See vbaretdecode.h. The one BSTR reader, shared with the arguments
    // column. `told` defaults false there: that column reaches this only
    // when the p-code named no type, so it is probing.
    bool DescribeBstrValue(std::uint64_t p, ValueWriter& w, bool told)
    {
        return DescribeBstr(p, w, told);
    }

    RetKind ExitReturnKind(std::uint16_t exitOp)
    {
        switch (exitOp)
        {
        case 623: return RetKind::Byte;
        case 624: return RetKind::Integer;
        case 625: return RetKind::Long;
        case 626: return RetKind::Single;
        case 627: return RetKind::Double;
        case 628: return RetKind::Currency;
        case 630: return RetKind::String;
        case 631: return RetKind::Object;
        case 634: return RetKind::LongLongOrArray;
        case 635: return RetKind::None;
        // 504 IS THE CLASS/FORM SUB EXIT. In a Sub [R14-8] is not a return slot,
        // it is a LOCAL, so leaving this unmapped read a local as a result.
        case 504: return RetKind::None;
        case 952: return RetKind::Variant;
        // 1494 IS THE RECORD FUNCTION EXIT. Unmapped, the store scan read a field of the
        // record as the whole result.
        case 1494: return RetKind::Record;
        // 1495 is the same for a record small enough to come back in a register.
        case 1495: return RetKind::RecordInFrame;
        default:  return RetKind::Unknown;
        }
    }

    const char* RetKindName(RetKind k)
    {
        switch (k)
        {
        case RetKind::Byte:     return "Byte";
        case RetKind::Integer:  return "Integer";
        case RetKind::Long:     return "Long";
        case RetKind::Single:   return "Single";
        case RetKind::Double:   return "Double";
        case RetKind::Currency: return "Currency";
        case RetKind::String:   return "String";
        case RetKind::Object:   return "Object";
        case RetKind::Variant:  return "Variant";
        case RetKind::Record:
        case RetKind::RecordInFrame: return "Udt";
        case RetKind::None:     return "Sub";
        default:                return "?";
        }
    }

    // THE KIND IS A PARAMETER, not re-derived here: class and form Functions
    // leave through opcode 1664 whatever they return, so their kind comes from
    // the STORE (DecideReturnKind), and deriving it in two places would be two
    // places to disagree.
    bool DescribeReturnKind(std::uint64_t r14, RetKind k, std::uint16_t storeOp,
                            std::int32_t exitOperand, ValueWriter& w, const char** typeOut)
    {
        *typeOut = "";
        if (r14 == 0) return false;

        switch (k)
        {
        case RetKind::None:
        case RetKind::Unknown:
            return false;

        // The record itself, by the address it is returned to: the same shape a record
        // argument takes, since its layout is not known here.
        case RetKind::Record:
        {
            std::uint64_t rec = 0;
            if (!RdU64(r14 + 8, rec) || rec == 0) return false;
            *typeOut = "Udt";
            w.Udt(rec);
            return true;
        }
        case RetKind::RecordInFrame:
        {
            if (exitOperand >= 0) return false;   // the record is a local, below R14
            *typeOut = "Udt";
            w.Udt(r14 + static_cast<std::int64_t>(exitOperand));
            return true;
        }

        case RetKind::String:
        {
            // A BSTR pointer, or NULL for the empty string -- VBA stores "" as
            // a null BSTR, and a null read as "" is the value, not a decline.
            std::uint64_t q = 0;
            if (!RdU64(r14 - 8, q)) return false;
            if (q == 0) { *typeOut = "String"; w.String(L"", 0, false); return true; }
            if (!DescribeBstr(q, w, /*told=*/true)) return false;
            *typeOut = "String";
            return true;
        }

        case RetKind::Variant:
            if (!DescribeVariantAt(r14 - 0x18, w, 0)) return false;
            *typeOut = "Variant";
            return true;

        case RetKind::Object:
        {
            // As Object / As Collection / As Range ...: the interface pointer
            // sits in the slot, or 0 for Nothing. Rendered as the
            // argument decoder renders an object.
            std::uint64_t q = 0;
            if (!RdU64(r14 - 8, q)) return false;
            *typeOut = "Object";
            DescribeObject(q, w);
            return true;
        }

        case RetKind::LongLongOrArray:
        {
            // Slot 634 serves both a LongLong and every typed array. The
            // opcode that STORED the slot tells them apart (699 = FStI8, 671 =
            // array assign); with neither in evidence the read is refused
            // rather than resolved by whether the bytes look like a pointer.
            std::uint64_t q = 0;
            if (!RdU64(r14 - 8, q)) return false;
            if (storeOp == 699)
            {
                *typeOut = "LongLong";
                PutNumber(w, "LongLong", "%lld", static_cast<long long>(static_cast<std::int64_t>(q)));
                return true;
            }
            if (storeOp == 671)
            {
                // THE ELEMENT TYPE COMES BACK FROM THE VALIDATED WALK, not from a read taken
                // before anything was validated. Passing 0 as the hint loses nothing: with
                // FADF_HAVEVARTYPE set the walk re-reads q-4 itself and ignores the hint.
                std::uint16_t vt = 0;
                // Never allocated: no descriptor, so nothing names the element type.
                if (q == 0) { w.Unallocated(nullptr); *typeOut = "?()"; return true; }
                if (!DescribeSafeArray(q, 0, w, 0, &vt))
                    return false;
                *typeOut = ArrayTypeName(vt);
                return true;
            }
            return false;
        }

        default:
        {
            std::uint64_t bits = 0;
            if (!RdU64(r14 - 8, bits)) return false;
            *typeOut = RetKindName(k);
            switch (k)
            {
            case RetKind::Byte:     return DescribeElement(17, r14 - 8, w);
            case RetKind::Integer:  return DescribeElement(2,  r14 - 8, w);
            case RetKind::Long:     return DescribeElement(3,  r14 - 8, w);
            case RetKind::Single:   return DescribeElement(4,  r14 - 8, w);
            case RetKind::Double:   return DescribeElement(5,  r14 - 8, w);
            case RetKind::Currency: return DescribeElement(6,  r14 - 8, w);
            default: return false;
            }
        }
        }
    }
}
