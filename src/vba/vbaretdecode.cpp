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

        using core::Append;

        // A BSTR, STRICTLY -- AND THIS IS THE ONLY ONE. The byte length sits in the
        // dword before the data, and everything here is a reason to refuse: a wrong
        // guess prints the first bytes of an unrelated allocation as though they
        // were a value. Truncation past kBstrMaxChars is marked; beyond
        // kBstrMaxBytes the whole thing is refused; a tab is text, other control
        // characters are the signature of reading memory that is not text.
        constexpr std::uint32_t kBstrMaxBytes = 4096;   // refuse beyond this when PROBING
        constexpr int           kBstrMaxChars = 256;    // then show this many

        // When the type is already known to be BSTR, long strings are allowed.
        constexpr std::uint32_t kBstrMaxBytesTold = 1u << 20;   // 512K characters

        // `told` -- does something ELSE already say this is a string? It decides
        // exactly one thing: whether a ZERO length is a value. When a VARIANT tag or
        // an exit opcode says BSTR, `cb == 0` is an empty string. When PROBING an
        // untyped slot, `cb == 0` is also what zeroed memory looks like, and is
        // refused.
        bool DescribeBstr(std::uint64_t p, char* out, int cap, bool told)
        {
            out[0] = 0;
            if (!InRangeAndAligned(p, 2)) return false;
            std::uint32_t cb = 0;
            if (!RdU32(p - 4, cb)) return false;
            if (cb & 1) return false;
            if (cb > (told ? kBstrMaxBytesTold : kBstrMaxBytes)) return false;
            const int n = static_cast<int>(cb / 2);
            if (n == 0)
            {
                if (!told) return false;          // zeroed memory looks like this
                _snprintf_s(out, cap, _TRUNCATE, "\"\"");
                return true;
            }

            wchar_t w[kBstrMaxChars + 2] = {};
            const int take = n < kBstrMaxChars ? n : kBstrMaxChars;
            if (!RdWide(p, w, take)) return false;

            // NUL-terminated PAST its full length -- checked at the real end,
            // which is why the byte limit and the character cap are separate.
            std::uint16_t term = 1;
            if (!RdU16(p + cb, term) || term != 0) return false;

            // A CONTROL CHARACTER IS A REFUSAL ONLY WHEN PROBING. Every other
            // check above is STRUCTURAL -- aligned, even length, bounded, and a
            // NUL at exactly `p + cb` -- and those are what actually tell a
            // string from a pointer or a double. This one looks at CONTENT, and
            // content proves nothing: a BSTR may legally hold any 16-bit unit.
            //
            // It earns its place only when nothing else has said "string": then
            // a byte below 32 is the best evidence available that the bits are
            // not text. When a VARIANT tag, an exit opcode or a declared
            // `As String` has ALREADY said BSTR, VBA's own metadata has settled
            // the type, and the same byte is a `vbCrLf` in a message or a tab in
            // a record. Refusing there threw a real value away and printed a raw
            // pointer where the string had been.
            if (!told)
                for (int i = 0; i < take; ++i)
                    if (w[i] < 32 && w[i] != 9) return false;

            return core::RenderQuoted(w, take, n > take, out, cap);
        }

        bool DescribeSafeArray(std::uint64_t psa, std::uint16_t vtHint, char* out, int cap,
                               int depth, int maxElems, std::uint16_t* vtOut = nullptr);
        bool DescribeVariantAt(std::uint64_t at, char* out, int cap, const char** typeOut,
                               int depth, int maxElems);

        // NESTING IS BOUNDED. It cannot be unbounded: each level is a real
        // recursion on the hot-path stack and the memory could be cyclic, so a
        // finite ceiling is the fail-safe. Deeper than this reads as "[...]"
        // rather than recursing, and the render buffer -- not this bound -- is
        // what stops a wide-but-shallow structure.
        //
        // THE COUNTER ADVANCES ON THE VARIANT-IN-A-VARIANT HOP, NOT ON THE
        // ARRAY: DescribeSafeArray passes `depth` through unchanged and only
        // DescribeElement's VT_VARIANT case increments.
        constexpr int kMaxNest = 32;

        // THE ONE PLACE AN OBJECT IS RENDERED, for both columns -- so an object
        // handed down as an argument and handed back as a result reads the same
        // on both rows.
        //
        // With OBJECTS on, the class is named and a Range, Worksheet or Workbook
        // is described. Anything that does not work out -- the setting off, a
        // class we cannot name, a COM call that failed -- falls back to the
        // ADDRESS, which is what this always said and is never wrong.
        volatile LONG g_describeObjects = 0;

        void DescribeObject(std::uint64_t ptr, char* out, int cap)
        {
            if (ptr == 0) { _snprintf_s(out, cap, _TRUNCATE, "Nothing"); return; }

            // THE ADDRESS IS ALWAYS THERE. The class replaces the word `object`
            // and the detail follows it; nothing is taken away. Dropping the
            // address was tried and caught by the test that follows ONE object
            // from an argument to a result -- which is what an address is for
            // and what a class name cannot do.
            if (InterlockedCompareExchange(&g_describeObjects, 0, 0) != 0)
            {
                // The detail is rendered to fit after "<class>@0x<16 hex>", so its
                // own "...(k of n shown)" marker is not what gets cut.
                char cls[128], detail[2048];
                constexpr int kHead = 32;
                const int detailCap = cap - kHead < static_cast<int>(sizeof(detail))
                                    ? cap - kHead : static_cast<int>(sizeof(detail));
                if (detailCap > 0 &&
                    DescribeObjectDetail(ptr, cls, sizeof(cls), detail, detailCap))
                {
                    const bool fits = strlen(cls) + 19 + strlen(detail) < static_cast<size_t>(cap);
                    _snprintf_s(out, cap, _TRUNCATE, "%s@0x%llX%s", cls,
                                static_cast<unsigned long long>(ptr), fits ? detail : "");
                    return;
                }
            }
            _snprintf_s(out, cap, _TRUNCATE, "object@0x%llX", static_cast<unsigned long long>(ptr));
        }

        // A DECIMAL: a 96-bit magnitude, a scale of 0..28 decimal places and a
        // sign byte, laid over the whole VARIANT (the vt shares its reserved
        // word): scale at +2, sign at +3, the high 32 bits at +4, the low 64 at
        // +8. Only ever inside a Variant -- `Decimal` cannot be declared, so
        // `CDec` is the one way one comes to exist. Rendered EXACTLY, in
        // integer arithmetic: 28 digits is more than a double holds, and a
        // rounded money figure is a wrong one. A scale above 28 or a sign byte
        // that is neither 0 nor 0x80 is not a DECIMAL, and is refused.
        bool DescribeDecimal(std::uint64_t at, char* out, int cap)
        {
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
            return true;
        }

        // One element of the given VARTYPE, read from `at`. `depth` is how
        // far inside nested arrays this element sits.
        bool DescribeElement(std::uint16_t vt, std::uint64_t at, char* out, int cap,
                             int depth = 0, int maxElems = kMaxRenderedElems)
        {
            out[0] = 0;
            std::uint64_t q = 0; std::uint32_t d = 0; std::uint16_t w = 0;
            switch (vt)
            {
            case 2:  // VT_I2 -- Integer AND Boolean share it
                if (!RdU16(at, w)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%d", static_cast<int>(static_cast<std::int16_t>(w)));
                return true;
            case 3:  // VT_I4
                if (!RdU32(at, d)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%d", static_cast<int>(static_cast<std::int32_t>(d)));
                return true;
            case 4:  // VT_R4
            {
                if (!RdU32(at, d)) return false;
                float f = 0; memcpy(&f, &d, sizeof(f));
                _snprintf_s(out, cap, _TRUNCATE, "%.9g", static_cast<double>(f));
                return true;
            }
            case 5:  // VT_R8
            case 7:  // VT_DATE -- a serial number, which is what a cell shows
            {
                if (!RdU64(at, q)) return false;
                double v = 0; memcpy(&v, &q, sizeof(v));
                // NaN and infinity are values; name them.
                if (((q >> 52) & 0x7FF) == 0x7FF)
                {
                    const bool neg  = (q >> 63) != 0;
                    const bool quiet = (q & 0x000FFFFFFFFFFFFFull) != 0;
                    _snprintf_s(out, cap, _TRUNCATE, "%s",
                                quiet ? "NaN" : (neg ? "-Infinity" : "Infinity"));
                    return true;
                }
                core::FormatDouble(v, out, cap);
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
                _snprintf_s(out, cap, _TRUNCATE, "%s%llu.%04llu",
                            neg ? "-" : "",
                            static_cast<unsigned long long>(a / 10000),
                            static_cast<unsigned long long>(a % 10000));
                return true;
            }
            case 8:  // VT_BSTR
                if (!RdU64(at, q)) return false;
                if (q == 0) { _snprintf_s(out, cap, _TRUNCATE, "\"\""); return true; }
                return DescribeBstr(q, out, cap, /*told=*/true);
            case 11: // VT_BOOL: VARIANT_TRUE is -1. Spelt as Excel spells it, like an error value.
                if (!RdU16(at, w)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%s", w ? "TRUE" : "FALSE");
                return true;
            case 17: // VT_UI1 -- one byte, read as one byte (see RdU8)
            {
                std::uint8_t b = 0;
                if (!RdU8(at, b)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%u", static_cast<unsigned>(b));
                return true;
            }
            case 20: // VT_I8
                if (!RdU64(at, q)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%lld", static_cast<long long>(static_cast<std::int64_t>(q)));
                return true;
            case 14: // VT_DECIMAL: in an array the element IS the whole DECIMAL
                return DescribeDecimal(at, out, cap);
            // THE FOREIGN INTEGER WIDTHS. No VBA declaration produces them, but a
            // COM property can hand one back inside a Variant, and a width is a
            // width: each is read at its size and printed as the number it is.
            case 16: // VT_I1
            {
                std::uint8_t b = 0;
                if (!RdU8(at, b)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%d", static_cast<int>(static_cast<std::int8_t>(b)));
                return true;
            }
            case 18: // VT_UI2
                if (!RdU16(at, w)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%u", static_cast<unsigned>(w));
                return true;
            case 19: case 23: // VT_UI4, VT_UINT
                if (!RdU32(at, d)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%u", static_cast<unsigned>(d));
                return true;
            case 22: // VT_INT
                if (!RdU32(at, d)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%d", static_cast<int>(static_cast<std::int32_t>(d)));
                return true;
            case 21: // VT_UI8
                if (!RdU64(at, q)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "%llu", static_cast<unsigned long long>(q));
                return true;
            case 9: case 13:  // VT_DISPATCH, VT_UNKNOWN
                if (!RdU64(at, q)) return false;
                DescribeObject(q, out, cap);
                return true;
            case 10: // VT_ERROR: a cell error travelling in an array (Range.Value)
            {
                if (!RdU32(at, d)) return false;
                // SPELT AS EXCEL SPELLS IT. `#N/A` is what the user sees in the
                // cell; `Error(0x7FA)` made them do the arithmetic. Unquoted, so
                // it cannot be confused with the string "#N/A", which renders
                // with its quotes.
                if (const char* en = core::ExcelErrNameFromVba(d))
                    _snprintf_s(out, cap, _TRUNCATE, "%s", en);
                else
                    _snprintf_s(out, cap, _TRUNCATE, "Error(0x%X)", d);
                return true;
            }
            case 12: // VT_VARIANT: the element IS a VARIANT, self-typed -- recurse
            {
                const char* ignored = "";
                return DescribeVariantAt(at, out, cap, &ignored, depth + 1, maxElems);
            }
            default:
            {
                // An element type this cannot read is named, not invented: a
                // placeholder says "there was one here, of type N" and leaves
                // the rest of the array intact.
                _snprintf_s(out, cap, _TRUNCATE, "?vt%u", static_cast<unsigned>(vt));
                return true;
            }
            }
        }


        // A SAFEARRAY, described as "<elem>[dims]{first,elements,...}" row by
        // row: the LAST index varies fastest, so a 2-D array reads a(0,0),
        // a(0,1), a(1,0) ... though the bytes lie column-major. Bounds come
        // from rgsabound, stored right-to-left (rgsabound[0] is the LAST
        // declared dimension), so they are printed in declaration order.
        //
        // Every header field is checked before anything is read through it:
        // a stale pointer that happens to be readable must not produce a
        // plausible-looking array out of unrelated memory.
        //
        // SIBLING of vbaargs.cpp's ReadSafeArray (the ARGUMENT path). Kept
        // separate on purpose -- this one takes a caller vtHint and is reached
        // directly, that one validates element width and follows an extra
        // indirection. See the fuller note there. Element rendering is shared
        // (DescribeArrayElement); the header parse is not, deliberately.
        // THE ONE ARRAY RENDERER, for both columns: "<Elem>[lo..hi,...]{e1,e2,...}".
        //
        // The header parse was merged into vbaoleaut.h and the ELEMENT decoder was
        // always shared -- this was the last piece written twice. The two copies
        // agreed on the truncation marker by hand and differed in their tail
        // reserve (32 against 48) and in how the element cap reached them, which
        // is how an array could read differently depending on which column it
        // landed in. The arguments column's `array[` fallback is gone with it:
        // an element type that cannot be named is `?`, the same `?` the type
        // position uses, meaning the same thing.
        //
        // `sa` must already have come from ReadSafeArrayHeader, which is what
        // makes the bounds and the element count safe to walk here.
        bool RenderSaText(const SaInfo& sa, std::uint16_t vt, char* out, int cap,
                          int depth, int maxElems)
        {
            const std::uint16_t cDims      = sa.cDims;
            const std::uint64_t total      = sa.total;
            const std::uint64_t pvData     = sa.pvData;
            const std::uint32_t cbElements = sa.cbElem;
            // BUILT IN THE CALLER'S BUFFER, WITH THE TAIL RESERVED. kTail bytes are held
            // back from every write below and the closing marker is written into that
            // reserve, so the elements can never squeeze it out. `cap` is the caller's,
            // so a NESTED array gets whatever room its parent had left and closes itself
            // on the same terms.
            constexpr int kTail = 48;            // ",...(64 of 18446744073709551615 shown)}"
            if (cap < kTail * 2) return false;   // no room to say anything true
            const int body = cap - kTail;        // what elements and bounds may use
            int len = 0;
            const char* en = VtName(vt);
            len = Append(out, body, len, en ? en : "?");
            len = Append(out, body, len, "[");
            // ALWAYS BOTH BOUNDS, because `Option Base` decides what a bare count
            // means -- and backwards, because rgsabound is stored in reverse of
            // VBA's declaration order.
            for (int d = static_cast<int>(cDims) - 1; d >= 0; --d)
            {
                char t[48];
                const long long lo = static_cast<long long>(sa.lBound[d]);
                _snprintf_s(t, _TRUNCATE, "%s%lld..%lld",
                            (d == static_cast<int>(cDims) - 1) ? "" : ",",
                            lo, lo + static_cast<long long>(sa.cElems[d]) - 1);
                len = Append(out, body, len, t);
            }
            len = Append(out, body, len, "]{");
            // HOW MANY ELEMENTS THE CALLER WANTS, not a constant.
            const std::uint64_t show = (maxElems > 0)
                                     ? static_cast<std::uint64_t>(maxElems)
                                     : static_cast<std::uint64_t>(kMaxRenderedElems);
            std::uint64_t shown = 0;
            // ROW BY ROW: the last declared index varies fastest. The bytes lie column-major
            // (the first index fastest), so each element's position in memory is computed.
            const int dims = cDims > 8 ? 8 : static_cast<int>(cDims);
            std::uint32_t extent[8] = {};
            std::uint64_t stride[8] = {};
            std::uint32_t idx[8] = {};
            for (int d = 0; d < dims; ++d)
            {
                extent[d] = sa.cElems[dims - 1 - d];               // declaration order
                stride[d] = d ? stride[d - 1] * extent[d - 1] : 1;
            }
            for (std::uint64_t i = 0; i < total && i < show; ++i)
            {
                // RENDER-IN-PLACE: the element is written straight into the
                // caller's buffer at the current offset, not into a fixed
                // 512-byte temp -- so a NESTED array or a long string element gets
                // whatever room is left in `body` rather than being cut at 512,
                // and nesting costs no per-level stack buffer. kMinElem guarantees
                // a whole SCALAR always fits before we start (the longest scalar
                // is ~25 chars), so a scalar is never left half-written; an
                // unbounded element fills the remaining room and, if it is a
                // nested array, marks its own truncation on the same terms.
                constexpr int kMinElem = 64;
                if (body - len < kMinElem) break;
                if (i) len = Append(out, body, len, ",");
                char* dst = out + len;
                const int dcap = body - len;      // room left, with the tail still reserved
                std::uint64_t at = 0;
                for (int d = 0; d < dims; ++d) at += idx[d] * stride[d];
                // ONE ELEMENT THE DECODER WILL NOT VOUCH FOR IS `?`, not a refusal of the
                // whole array: the shape and size were established from the descriptor and
                // stay true whatever a single element turns out to hold.
                if (!DescribeElement(vt, pvData + at * cbElements, dst, dcap, depth, maxElems))
                    _snprintf_s(dst, dcap, _TRUNCATE, "?");
                len += static_cast<int>(strlen(dst));
                ++shown;
                for (int d = dims - 1; d >= 0; --d) { if (++idx[d] < extent[d]) break; idx[d] = 0; }
            }
            // AGAINST WHAT WAS SHOWN, covering BOTH reasons for stopping -- the element
            // cap and the buffer -- in the words the arguments column uses.
            if (shown < total)
            {
                len = core::AppendShownMarker(out, cap, len, shown, total);
            }
            len = Append(out, cap, len, "}");
            return true;
        }


        bool DescribeSafeArray(std::uint64_t psa, std::uint16_t vtHint, char* out, int cap,
                               int depth, int maxElems, std::uint16_t* vtOut)
        {
            if (vtOut) *vtOut = 0;
            if (depth > kMaxNest) { _snprintf_s(out, cap, _TRUNCATE, "[...]"); return true; }
            out[0] = 0;

            // One SAFEARRAY reader for both columns (vbaoleaut.h).
            SaInfo sa{};
            if (!ReadSafeArrayHeader(psa, vtHint, sa)) return false;
            const std::uint16_t vt = EffectiveElemVt(sa);
            // REPORTED BACK, so a caller that needs to NAME the element type does not
            // read the descriptor a second time.
            if (vtOut) *vtOut = vt;
            return RenderSaText(sa, vt, out, cap, depth, maxElems);
        }

        // Is this a VARTYPE the decoder can actually read a value for?
        //
        // ONLY CONSULTED AT THE TOP LEVEL OF A RESULT, and the asymmetry is the
        // point. INSIDE an array an unreadable element renders `?vtN` and the
        // array survives, because the other elements are real. As the WHOLE
        // result there is nothing else in the row, so `?vt41231` would be the
        // entire answer -- a claim that a Variant was returned, made from bytes
        // whose own tag says they are not one.
        //
        // Reachable because a Variant is identified by its SLOT: anything
        // at [R14-0x18] is read as one. If a procedure ever puts something else
        // there the row goes empty and is COUNTED, rather than confident and
        // wrong.
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

        // A VARIANT at `at`: vt, then the value at +8, decoded by vt.
        bool DescribeVariantAt(std::uint64_t at, char* out, int cap, const char** typeOut,
                               int depth, int maxElems)
        {
            out[0] = 0;
            std::uint16_t vt = 0;
            if (!RdU16(at, vt)) return false;
            std::uint64_t val  = at + 8;          // the payload
            std::uint64_t self = at;              // where a DECIMAL overlays the VARIANT
            // True after a by-reference hop; VT_RECORD needs to know.
            bool byref = false;

            // VT_BYREF: the VARIANT holds a POINTER to the value. That is how
            // the interpreter passes a typed variable into a Variant parameter
            // -- 0x4008 is VT_BYREF|VT_BSTR -- so the shape is
            // ordinary VBA, not a corner. Followed once, under the same guard
            // as every other pointer, and at the top level only: inside an
            // array a by-reference element is not something the array owns.
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

            if (vt & kVT_ARRAY)                   // the element type is in the low bits
            {
                std::uint64_t psa = 0;
                if (!RdU64(val, psa)) return false;
                *typeOut = "Variant";
                return DescribeSafeArray(psa, static_cast<std::uint16_t>(vt & kVT_TYPEMASK), out, cap,
                                         depth, maxElems);
            }
            if (depth == 0 && !ReadableVartype(vt)) return false;   // see ReadableVartype

            *typeOut = "Variant";
            switch (vt)
            {
            case 0:  _snprintf_s(out, cap, _TRUNCATE, "Empty"); return true;
            case 1:  _snprintf_s(out, cap, _TRUNCATE, "Null");  return true;
            case 10: // VT_ERROR: the SCODE, which for a UDF is a cell error
            {
                std::uint32_t sc = 0;
                if (!RdU32(val, sc)) return false;
                // The same spelling the XLL column uses, from one table: a `#N/A`
                // an XLL returned and a `#N/A` a UDF was handed are the same fact
                // about the same sheet. An SCODE outside Excel's set keeps its
                // number -- a COM error is not a cell error, and naming it would
                // claim it was.
                if (const char* en = core::ExcelErrNameFromVba(sc))
                    _snprintf_s(out, cap, _TRUNCATE, "%s", en);
                else
                    _snprintf_s(out, cap, _TRUNCATE, "Error(0x%X)", sc);
                return true;
            }
            case 9: case 13:                       // objects
            {
                std::uint64_t ptr = 0;
                if (!RdU64(val, ptr)) return false;
                DescribeObject(ptr, out, cap);
                return true;
            }
            case 14:                               // the DECIMAL overlays the VARIANT itself
                return DescribeDecimal(self, out, cap);
            case 36:                               // VT_RECORD: a user-defined Type
            {
                // Its layout lives behind IRecordInfo, a COM call this hook does
                // not make, so it is its address -- the same shape the argument
                // column gives a UDT, so one can be followed between rows.
                //
                // After a by-reference hop, val already is the record's address.
                std::uint64_t rec = val;
                if (!byref && !RdU64(val, rec)) return false;
                _snprintf_s(out, cap, _TRUNCATE, "udt@0x%llX", static_cast<unsigned long long>(rec));
                return true;
            }
            default:
                return DescribeElement(vt, val, out, cap, depth);
            }
        }
    }

    // THE TYPE FROM THE STORE, when the exit opcode does not carry one.
    //
    // Class and form Functions all leave through exit opcode 1664 whatever they
    // return -- measured across eleven procedures and every declared type -- so
    // the instruction that WROTE the result carries the type, exactly as it does
    // for slot 634. The store family is the load family plus 32, measured on
    // seven types (vbapcode.h). String, Object and Variant do NOT follow that
    // pairing and were each measured in isolation.
    //
    // Boolean needs no entry: VBA stores it as an Integer (-1/0).
    //
    // NOTE THE ASYMMETRY WITH VARIANT. 694 is here, but it is not what identifies
    // a Variant -- the caller does that from the OPERAND (-0x18), because a
    // Variant is stored through a different opcode depending on what it holds
    // (707 for a string). See DecideReturnKind in vbatrace.cpp.
    RetKind StoreReturnKind(std::uint16_t storeOp)
    {
        switch (storeOp)
        {
        case 688: return RetKind::Byte;
        case 689: return RetKind::Integer;
        case 690: return RetKind::Long;
        case 693: return RetKind::Currency;
        case 699: return RetKind::LongLongOrArray;
        // 671 IS THE TYPED-ARRAY STORE, here for the same reason 699 is: on the
        // CLASS and FORM path the exit opcode is 1664, which carries no type, so
        // the store is the only thing that can name the result.
        //
        // NOT SUFFICIENT ON ITS OWN, AND MEASURED SO: with this in place a class
        // `Function RLongArr() As Long()` STILL reports an empty ret and 1664 is
        // still counted unmapped. So either the store scan is not finding 671 at
        // operand -8 for this shape or it declines on a conflict -- the next
        // thing to MEASURE, not to assume. The row stays because it is right
        // about what 671 means; the defect is not closed.
        // [defect: class-function-array-return-unmapped]
        case 671: return RetKind::LongLongOrArray;
        case 700: return RetKind::Single;
        case 701: return RetKind::Double;

        // ---- STRING, VARIANT AND OBJECT ---------------------------------
        //
        // Each cross-checks against where the decoder already reads it:
        //   708 stores at operand -8     -- where String reads its BSTR
        //   694 stores at operand -0x18  -- where Variant reads, and nowhere else does
        //   696 stores at operand -8     -- and 696-32 = 664, the load table's Object
        //
        // A Variant holding an array uses the same store as a scalar Variant
        // (measured: `Array(1,2,3)` and `42` both store through 694@-0x18); one
        // holding a string stores through 707 at the same -0x18 -- which is why the
        // caller decides a Variant by its OPERAND and not by this table. Nothing here
        // distinguishes the held type; DescribeVariantAt reads the VARIANT's own tag.
        case 708: return RetKind::String;
        case 694: return RetKind::Variant;
        case 696: return RetKind::Object;

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
    bool RenderSafeArrayValue(const SaInfo& sa, char* out, int cap, int maxElems)
    {
        return RenderSaText(sa, EffectiveElemVt(sa), out, cap, 0, maxElems);
    }

    // See vbaretdecode.h. One element of a SAFEARRAY, for a caller that walked
    // the descriptor itself.
    bool DescribeArrayElement(std::uint16_t vt, std::uint64_t at, char* out, int cap, int maxElems)
    {
        out[0] = 0;
        if (vt == 0) return false;
        return DescribeElement(vt, at, out, cap, 0, maxElems);
    }

    // See vbaretdecode.h. The ARGS column's decoder, which is this one.
    bool DescribeVariantValue(std::uint64_t at, char* out, int cap,
                              const char** heldType, int maxElems)
    {
        *heldType = "";
        out[0] = 0;

        // The held type is read BEFORE the value, from the same tag the
        // renderer will use, so the two cannot disagree about what was found.
        std::uint16_t vt = 0;
        if (!RdU16(at, vt)) return false;

        // An ARRAY names its own element type in the rendering
        // (`Double[3]{...}`), so naming it again would read `Double(Double[3]
        // {...})`. Same for the cases whose text already carries the type, and
        // for the ones with no type worth saying.
        if ((vt & kVT_ARRAY) == 0)
        {
            // Through the VT_BYREF tag: the name is the target's, which is what
            // the value will be read as.
            // ONE NAME TABLE. This was a hand-copied SUBSET of VtName -- the same
            // sixteen strings, maintained separately -- so renaming a type in one
            // place would have made the two columns disagree about it. The subset
            // is now a policy (`VtNameWorthSaying`) beside the fact (`VtName`).
            const std::uint16_t base = static_cast<std::uint16_t>(vt & ~kVT_BYREF);
            if (VtNameWorthSaying(base)) *heldType = VtName(base);
        }

        const char* ignored = "";
        if (!DescribeVariantAt(at, out, cap, &ignored, 0, maxElems))
        {
            *heldType = "";   // nothing was rendered, so nothing is named
            out[0] = 0;
            return false;
        }
        return true;
    }

    // See vbaretdecode.h. The one BSTR reader, shared with the arguments
    // column. `told` defaults false there: that column reaches this only
    // when the p-code named no type, so it is probing.
    bool DescribeBstrValue(std::uint64_t p, char* out, int cap, bool told)
    {
        return DescribeBstr(p, out, cap, told);
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
        case RetKind::None:     return "Sub";
        default:                return "?";
        }
    }

    // THE KIND IS A PARAMETER, not re-derived here: class and form Functions
    // leave through opcode 1664 whatever they return, so their kind comes from
    // the STORE (DecideReturnKind), and deriving it in two places would be two
    // places to disagree.
    bool DescribeReturnKind(std::uint64_t r14, RetKind k, std::uint16_t storeOp,
                            char* out, int cap, const char** typeOut)
    {
        out[0] = 0; *typeOut = "";
        if (r14 == 0) return false;

        switch (k)
        {
        case RetKind::None:
        case RetKind::Unknown:
            return false;

        case RetKind::String:
        {
            // A BSTR pointer, or NULL for the empty string -- VBA stores "" as
            // a null BSTR, and a null read as "" is the value, not a decline.
            std::uint64_t q = 0;
            if (!RdU64(r14 - 8, q)) return false;
            *typeOut = "String";
            if (q == 0) { _snprintf_s(out, cap, _TRUNCATE, "\"\""); return true; }
            return DescribeBstr(q, out, cap, /*told=*/true);
        }

        case RetKind::Variant:
            return DescribeVariantAt(r14 - 0x18, out, cap, typeOut, 0, kMaxRenderedElems);

        case RetKind::Object:
        {
            // As Object / As Collection / As Range ...: the interface pointer
            // sits in the slot, or 0 for Nothing. Rendered as the
            // argument decoder renders an object.
            std::uint64_t q = 0;
            if (!RdU64(r14 - 8, q)) return false;
            *typeOut = "Object";
            DescribeObject(q, out, cap);
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
                _snprintf_s(out, cap, _TRUNCATE, "%lld", static_cast<long long>(static_cast<std::int64_t>(q)));
                return true;
            }
            if (storeOp == 671)
            {
                // THE ELEMENT TYPE COMES BACK FROM THE VALIDATED WALK, not from a read taken
                // before anything was validated. Passing 0 as the hint loses nothing: with
                // FADF_HAVEVARTYPE set the walk re-reads q-4 itself and ignores the hint.
                std::uint16_t vt = 0;
                if (!DescribeSafeArray(q, 0, out, cap, 0, kMaxRenderedElems, &vt))
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
            case RetKind::Byte:     return DescribeElement(17, r14 - 8, out, cap);
            case RetKind::Integer:  return DescribeElement(2,  r14 - 8, out, cap);
            case RetKind::Long:     return DescribeElement(3,  r14 - 8, out, cap);
            case RetKind::Single:   return DescribeElement(4,  r14 - 8, out, cap);
            case RetKind::Double:   return DescribeElement(5,  r14 - 8, out, cap);
            case RetKind::Currency: return DescribeElement(6,  r14 - 8, out, cap);
            default: return false;
            }
        }
        }
    }
}
