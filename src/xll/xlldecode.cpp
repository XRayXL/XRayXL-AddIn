#include "xlldecode.h"
#include "core/excelerr.h"
#include "xlcall.h"
#include "core/safemem.h"
#include "core/text.h"
#include "core/xltype.h"

#include <cstddef>
#include <cstdio>
#include <cstring>

namespace xll
{
    // Every pointer here came from somebody else's add-in: the guarded read
    // is core's, and a fault costs the value, not the process.
    static bool SafeRead(void* dst, const void* src, size_t n)
    {
        return src != nullptr && dst != nullptr &&
               core::RdBytes(reinterpret_cast<std::uint64_t>(src), dst, n);
    }

    namespace
    {
        using core::ValueWriter;

        void PutInt(ValueWriter& w, const char* type, long long v)
        {
            char num[24]; _snprintf_s(num, _TRUNCATE, "%lld", v);
            w.Number(type, num);
        }

        // A scalar behind a pointer that came from somebody else's add-in.
        template <class T>
        bool ReadAt(ULONG64 p, T& v)
        {
            return SafeRead(&v, reinterpret_cast<const void*>(p), sizeof(T));
        }

        // ---- strings: four conventions, each with its own reader ------------
        // Only the registered code tells them apart; no length check can.

        // Reads up to n bytes, stopping at the page boundary first, so a short
        // string that ends just before unmapped memory is still read. Returns bytes read.
        size_t SafeReadString(void* dst, const void* src, size_t n)
        {
            const size_t addr = reinterpret_cast<size_t>(src);
            size_t first = 0x1000 - (addr & 0xFFF);
            if (first > n) first = n;
            if (!SafeRead(dst, src, first)) return 0;
            if (first == n) return n;
            return SafeRead(static_cast<char*>(dst) + first,
                            static_cast<const char*>(src) + first, n - first) ? n : first;
        }

        // Narrow strings arrive in the ANSI code page.
        void FinishAnsi(ValueWriter& w, const char* text, int len, bool cut)
        {
            wchar_t wide[kValueMax];
            const int m = len > 0 ? MultiByteToWideChar(CP_ACP, 0, text, len, wide, kValueMax) : 0;
            w.String(wide, m > 0 ? m : 0, cut);
        }

        bool ReadAsciiZ(const void* p, ValueWriter& w)
        {
            char buf[kValueMax];
            const size_t got = SafeReadString(buf, p, sizeof(buf));
            if (got == 0) return false;
            // Cut only when no NUL was seen in what was read.
            const int avail = static_cast<int>(got < sizeof(buf) ? got : sizeof(buf));
            int len = 0;
            while (len < avail && buf[len] != 0) len++;
            const bool cut = (len == avail);
            if (len > static_cast<int>(sizeof(buf)) - 1) len = static_cast<int>(sizeof(buf)) - 1;
            FinishAnsi(w, buf, len, cut);
            return true;
        }

        bool ReadAsciiCounted(const void* p, ValueWriter& w)
        {
            unsigned char n = 0;
            if (!SafeRead(&n, p, 1)) return false;
            int len = n;
            const bool cut = len > kValueMax;
            if (cut) len = kValueMax;
            char buf[256];
            if (len > 0 && !SafeRead(buf, static_cast<const unsigned char*>(p) + 1, len)) return false;
            FinishAnsi(w, buf, len, cut);
            return true;
        }

        bool ReadWideZ(const void* p, ValueWriter& w)
        {
            wchar_t wbuf[kValueMax];
            const size_t got = SafeReadString(wbuf, p, sizeof(wbuf)) / sizeof(wchar_t);
            if (got == 0) return false;
            // Cut only when no NUL was seen in what was read.
            const int avail = static_cast<int>(got < kValueMax ? got : kValueMax);
            int len = 0;
            while (len < avail && wbuf[len] != 0) len++;
            const bool cut = (len == avail);
            if (len > kValueMax - 1) len = kValueMax - 1;
            w.String(wbuf, len, cut);
            return true;
        }

        bool ReadWideCounted(const void* p, ValueWriter& w)
        {
            wchar_t n = 0;
            if (!SafeRead(&n, p, sizeof(n))) return false;
            const int real = static_cast<int>(n);   // a WCHAR count: 0..32767
            if (real < 0 || real > 32767) return false;
            int len = real;
            const bool cut = len > kValueMax;
            if (cut) len = kValueMax;
            wchar_t wbuf[kValueMax];
            if (len > 0 && !SafeRead(wbuf, static_cast<const wchar_t*>(p) + 1, len * sizeof(wchar_t)))
                return false;
            w.String(wbuf, len, cut);
            return true;
        }

        // Spellings shared with the VBA column; an unknown code prints `#ERR?` to fit the fixed-width grid.
        const char* ErrName(int e)
        {
            const char* n = core::ExcelErrName(e);
            return n ? n : "#ERR?";
        }

        // Anything larger than Excel's grid is not an array Excel made, so it is refused, not walked.
        bool GridShapeOk(long long rows, long long cols)
        {
            return rows >= 0 && cols >= 0 && rows <= 1048576 && cols <= 16384;
        }

        // Shared by FP, FP12, the type-O triple and XLOPER arrays; Excel stores them row-major.
        // `cell` returns false when it cannot read element i, which then reads `?`.
        template <class CellFn>
        bool WriteGrid(ValueWriter& w, const char* elem, long long rows, long long cols, CellFn cell)
        {
            if (!GridShapeOk(rows, cols)) return false;
            const long long lo[2] = { 1, 1 };
            const long long hi[2] = { rows, cols };
            const std::uint64_t extent[2] = { static_cast<std::uint64_t>(rows),
                                              static_cast<std::uint64_t>(cols) };
            w.BeginArray(elem, 2, lo, hi);
            core::WalkRowMajor(w, 2, extent, [&](const std::uint64_t* idx)
            {
                const long long i = static_cast<long long>(idx[0]) * cols + static_cast<long long>(idx[1]);
                const ValueWriter::Mark m = w.Save();
                if (!cell(i)) { w.Restore(m); w.Marker("?"); }
            });
            w.EndArray();
            return true;
        }

        // A run of doubles, which is what FP, FP12 and the type-O triple all
        // hold once their headers have been read.
        bool WriteDoubleGrid(ValueWriter& w, const double* cells, long long rows, long long cols)
        {
            return WriteGrid(w, "Double", rows, cols, [&](long long i)
            {
                double d = 0;
                if (!SafeRead(&d, cells + i, sizeof(d))) return false;
                w.Double(d);
                return true;
            });
        }

        // Copy only the header; the doubles are read one at a time.
        template <class Fp>
        bool ReadFp(const void* p, ValueWriter& w)
        {
            Fp h{};
            if (!SafeRead(&h, p, offsetof(Fp, array))) return false;
            const double* cells = reinterpret_cast<const double*>(
                static_cast<const unsigned char*>(p) + offsetof(Fp, array));
            return WriteDoubleGrid(w, cells, h.rows, h.columns);
        }

        // Handle both the small grid XLOPER and the big grid XLOPER12. xltypeInt is 16 bits in
        // one and 32 in the other, named as VBA names those widths.
        struct NarrowOper
        {
            using Oper = XLOPER;    using Ref = XLREF;    using MRef = XLMREF;
            static constexpr const char* kInt = "Integer";
            static bool Str(const void* s, ValueWriter& w) { return ReadAsciiCounted(s, w); }
        };
        struct WideOper
        {
            using Oper = XLOPER12;  using Ref = XLREF12;  using MRef = XLMREF12;
            static constexpr const char* kInt = "Long";
            static bool Str(const void* s, ValueWriter& w) { return ReadWideCounted(s, w); }
        };

        // One area, 1-based.
        template <class Ref>
        void WriteArea(ValueWriter& w, const Ref& a)
        {
            w.Area(static_cast<long long>(a.rwFirst) + 1, static_cast<long long>(a.colFirst) + 1,
                   static_cast<long long>(a.rwLast) + 1, static_cast<long long>(a.colLast) + 1);
        }

        template <class T>
        bool DescribeOper(const void* p, ValueWriter& w, int depth = 0);

        template <class T>
        bool OperBody(const typename T::Oper& x, ValueWriter& w, int depth)
        {
            const int t = x.xltype & core::kXlTypeMask;
            switch (t)
            {
            case xltypeNum:     w.Double(x.val.num);                             return true;
            case xltypeStr:     return T::Str(x.val.str, w);
            case xltypeBool:    w.Bool(x.val.xbool != 0);                        return true;
            case xltypeErr:     w.Error(ErrName(static_cast<int>(x.val.err)));   return true;
            case xltypeInt:     PutInt(w, T::kInt, x.val.w);                     return true;
            case xltypeMissing: w.Word("Missing");                               return true;
            case xltypeNil:     w.Word("Empty");                                 return true;
            case xltypeSRef:
                w.BeginReference("SRef");
                WriteArea(w, x.val.sref.ref);
                w.EndReference();
                return true;
            case xltypeRef:
            {
                const void* m = x.val.mref.lpmref;
                using MRef = typename T::MRef;
                using Ref = typename T::Ref;
                MRef head{};
                if (!SafeRead(&head, m, offsetof(MRef, reftbl))) return false;
                w.BeginReference("Ref");
                for (unsigned i = 0; i < head.count; ++i)
                {
                    Ref area{};
                    const void* at = static_cast<const unsigned char*>(m) + offsetof(MRef, reftbl) + i * sizeof(Ref);
                    if (!SafeRead(&area, at, sizeof(area))) return false;
                    WriteArea(w, area);
                }
                w.EndReference();
                return true;
            }
            case xltypeMulti:
            {
                const unsigned char* arr = static_cast<const unsigned char*>(
                    static_cast<const void*>(x.val.array.lparray));
                return WriteGrid(w, "Variant", x.val.array.rows, x.val.array.columns, [&](long long i)
                {
                    return DescribeOper<T>(arr + i * sizeof(typename T::Oper), w, depth + 1);
                });
            }
            default:
            {
                char tmp[32]; _snprintf_s(tmp, _TRUNCATE, "?xltype%d", t);
                w.Marker(tmp);
                return true;
            }
            }
        }

        // An XLOPER is Excel's variant, so its value is written as a Variant's is. False,
        // having written nothing, when it cannot be read.
        template <class T>
        bool DescribeOper(const void* p, ValueWriter& w, int depth)
        {
            // An array element that is itself an array is not followed further.
            if (depth > 1) { w.Marker("[...]"); return true; }
            typename T::Oper x{};
            if (!SafeRead(&x, p, sizeof(x))) return false;
            const ValueWriter::Mark m = w.Save();
            w.BeginVariant();
            if (!OperBody<T>(x, w, depth)) { w.Restore(m); return false; }
            w.EndVariant();
            return true;
        }

        // One value, whatever slot it came from: `bits` is the integer register or stack slot,
        // `dbl` the XMM one. Writes nothing when the value could not be established.
        void DescribeValue(Kind k, ULONG64 bits, double dbl, ValueWriter& w)
        {
            const void* p = reinterpret_cast<const void*>(bits);
            switch (k)
            {
            case Kind::Double:    w.Double(dbl);                                                  return;
            case Kind::Int16:     PutInt(w, "Integer", static_cast<short>(bits & 0xFFFF));         return;
            case Kind::UInt16:    PutInt(w, "UInt16", static_cast<unsigned short>(bits & 0xFFFF)); return;
            case Kind::Int32:     PutInt(w, "Long", static_cast<INT32>(bits & 0xFFFFFFFF));        return;
            case Kind::PtrInt16:  { short  v = 0; if (ReadAt(bits, v)) PutInt(w, "Integer", v); return; }
            case Kind::PtrInt32:  { INT32  v = 0; if (ReadAt(bits, v)) PutInt(w, "Long", v);    return; }
            case Kind::PtrDouble: { double v = 0; if (ReadAt(bits, v)) w.Double(v);              return; }
            case Kind::StrAsciiZ:     ReadAsciiZ(p, w);                return;
            case Kind::StrAsciiCount: ReadAsciiCounted(p, w);          return;
            case Kind::StrWideZ:      ReadWideZ(p, w);                 return;
            case Kind::StrWideCount:  ReadWideCounted(p, w);           return;
            case Kind::Fp:            ReadFp<FP>(p, w);                return;
            case Kind::Fp12:          ReadFp<FP12>(p, w);              return;
            case Kind::OperNarrow:    DescribeOper<NarrowOper>(p, w);  return;
            case Kind::OperWide:      DescribeOper<WideOper>(p, w);    return;
            default:                  return;
            }
        }
    }

    // ---- one argument ------------------------------------------------------
    void DescribeArg(const Slot& s, const Regs& r, core::ValueWriter& w)
    {
        const int pos = s.abiIndex;

        switch (s.kind)
        {
        case Kind::ArrayTriple:
        {
            // Three slots: the head renders the whole array, the other two write nothing.
            if (!s.tripleHead) return;
            // xlcall.h has no struct for O: rows*, columns*, then the doubles.
            long long rows = 0, cols = 0;
            if (s.code[1] == '%')
            {
                INT32 v = 0;
                if (ReadAt(IntArgAt(r, pos), v))     rows = v;
                if (ReadAt(IntArgAt(r, pos + 1), v)) cols = v;
            }
            else
            {
                WORD v = 0;
                if (ReadAt(IntArgAt(r, pos), v))     rows = v;
                if (ReadAt(IntArgAt(r, pos + 1), v)) cols = v;
            }
            WriteDoubleGrid(w, reinterpret_cast<const double*>(IntArgAt(r, pos + 2)), rows, cols);
            return;
        }
        case Kind::AsyncHandle:
            w.Word("AsyncHandle");
            return;
        default:
            // Position picks the index, type picks the register file (xllregs.h).
            DescribeValue(s.kind, IntArgAt(r, pos), DoubleArgAt(r, pos), w);
            return;
        }
    }

    // ---- the return value --------------------------------------------------
    void DescribeReturn(Kind k, const Regs& r, core::ValueWriter& w)
    {
        switch (k)
        {
        // Nothing to report: an async call answers later, and a void or in-place return has no value.
        case Kind::AsyncHandle: return;
        case Kind::Void:        return;
        default:
            // A Double is in XMM0; rax holds a leftover that would decode as a plausible pointer.
            DescribeValue(k, r.rax, r.xmmRet, w);
            return;
        }
    }
}
