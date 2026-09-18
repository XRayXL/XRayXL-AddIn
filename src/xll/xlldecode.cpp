#include "xlldecode.h"
#include "core/excelerr.h"
#include "xlcall.h"
#include "core/render.h"
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
        // Copies `text`, ending "..." if it had to be cut.
        void Put(char* out, int outSize, const char* text)
        {
            if (outSize <= 0) return;
            int i = 0;
            for (; text[i] != 0 && i < outSize - 1; i++) out[i] = text[i];
            out[i] = 0;
            if (text[i] != 0 && outSize >= 5) memcpy(out + outSize - 4, "...", 4);
        }

        using core::Append;

        using core::FormatDouble;      // the VBA column's, so the two agree
        void PutDouble(char* out, int outSize, double d)
        {
            char num[48]; FormatDouble(d, num, sizeof(num));
            Put(out, outSize, num);
        }
        void PutInt(char* out, int outSize, long long v)
        {
            char num[24]; _snprintf_s(num, _TRUNCATE, "%lld", v);
            Put(out, outSize, num);
        }

        // A scalar behind a pointer that came from somebody else's add-in.
        template <class T>
        bool ReadAt(ULONG64 p, T& v)
        {
            return SafeRead(&v, reinterpret_cast<const void*>(p), sizeof(T));
        }

        // ---- strings: four conventions, each with its own reader ------------
        //
        // Only the registered code tells them apart; no length check can. All
        // four are rendered with the VBA column's quoting (core::RenderQuoted).

        // The same element cap the VBA column uses.
        const long long kMaxElems = core::kMaxRenderedElems;

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

        void FinishWide(char* out, int outSize, const wchar_t* w, int len, bool cut)
        {
            if (!core::RenderQuoted(w, len, cut, out, outSize)) Put(out, outSize, "");
        }

        // Narrow strings arrive in the ANSI code page.
        void FinishAnsi(char* out, int outSize, const char* text, int len, bool cut)
        {
            wchar_t w[kValueMax];
            const int m = len > 0 ? MultiByteToWideChar(CP_ACP, 0, text, len, w, kValueMax) : 0;
            FinishWide(out, outSize, w, m > 0 ? m : 0, cut);
        }

        void ReadAsciiZ(const void* p, char* out, int outSize)
        {
            char buf[kValueMax];
            const size_t got = SafeReadString(buf, p, sizeof(buf));
            if (got == 0) { Put(out, outSize, ""); return; }
            // Cut only when no NUL was seen in what was read.
            const int avail = static_cast<int>(got < sizeof(buf) ? got : sizeof(buf));
            int len = 0;
            while (len < avail && buf[len] != 0) len++;
            const bool cut = (len == avail);
            if (len > static_cast<int>(sizeof(buf)) - 1) len = static_cast<int>(sizeof(buf)) - 1;
            FinishAnsi(out, outSize, buf, len, cut);
        }

        void ReadAsciiCounted(const void* p, char* out, int outSize)
        {
            unsigned char n = 0;
            if (!SafeRead(&n, p, 1)) { Put(out, outSize, ""); return; }
            int len = n;
            const bool cut = len > kValueMax;
            if (cut) len = kValueMax;
            char buf[256];
            if (len > 0 && !SafeRead(buf, static_cast<const unsigned char*>(p) + 1, len))
            { Put(out, outSize, ""); return; }
            FinishAnsi(out, outSize, buf, len, cut);
        }

        void ReadWideZ(const void* p, char* out, int outSize)
        {
            wchar_t wbuf[kValueMax];
            const size_t got = SafeReadString(wbuf, p, sizeof(wbuf)) / sizeof(wchar_t);
            if (got == 0) { Put(out, outSize, ""); return; }
            // Cut only when no NUL was seen in what was read.
            const int avail = static_cast<int>(got < kValueMax ? got : kValueMax);
            int len = 0;
            while (len < avail && wbuf[len] != 0) len++;
            const bool cut = (len == avail);
            if (len > kValueMax - 1) len = kValueMax - 1;
            FinishWide(out, outSize, wbuf, len, cut);
        }

        void ReadWideCounted(const void* p, char* out, int outSize)
        {
            wchar_t n = 0;
            if (!SafeRead(&n, p, sizeof(n))) { Put(out, outSize, ""); return; }
            const int real = static_cast<int>(n);   // a WCHAR count: 0..32767
            if (real < 0 || real > 32767) { Put(out, outSize, ""); return; }
            int len = real;
            const bool cut = len > kValueMax;
            if (cut) len = kValueMax;
            wchar_t wbuf[kValueMax];
            if (len > 0 && !SafeRead(wbuf, static_cast<const wchar_t*>(p) + 1, len * sizeof(wchar_t)))
            { Put(out, outSize, ""); return; }
            FinishWide(out, outSize, wbuf, len, cut);
        }

        // Spellings shared with the VBA column (core/excelerr.h). `#ERR?` stays here: this
        // column must print something in a fixed-width grid, where the VBA column can print the
        // number.
        const char* ErrName(int e)
        {
            const char* n = core::ExcelErrName(e);
            return n ? n : "#ERR?";
        }

        // One grid renderer: "<Elem>[1..<rows>,1..<cols>]{a,b,...}", row by row, for an FP, an
        // FP12, a type-O triple and an XLOPER array alike. `cell` returns false when it cannot
        // render element i, which ends the grid.
        template <class CellFn>
        void RenderGrid(const char* elem, long long rows, long long cols, char* out, int outSize, CellFn cell)
        {
            if (rows < 0 || cols < 0 || rows > 1000000 || cols > 1000000)
            { Put(out, outSize, ""); return; }
            // Room is held back for the closing marker, as the VBA column does,
            // so a full buffer still ends with an honest count.
            char tmp[kValueMax];
            constexpr int kTail = 32;
            const int body = static_cast<int>(sizeof(tmp)) - kTail;
            int len = 0;
            char head[64];
            _snprintf_s(head, _TRUNCATE, "%s[1..%lld,1..%lld]{", elem, rows, cols);
            len = Append(tmp, body, len, head);
            const long long total = rows * cols;
            long long shown = 0;
            for (long long i = 0; i < total && i < kMaxElems; i++)
            {
                char c[64] = { 0 };
                if (!cell(i, c, static_cast<int>(sizeof(c)))) break;
                const int need = static_cast<int>(strlen(c)) + (i ? 1 : 0);
                if (len + need >= body) break;
                if (i) len = Append(tmp, body, len, ",");
                len = Append(tmp, body, len, c);
                ++shown;
            }
            if (shown < total)
                len = core::AppendShownMarker(tmp, sizeof(tmp), len,
                                              static_cast<unsigned long long>(shown),
                                              static_cast<unsigned long long>(total));
            Append(tmp, sizeof(tmp), len, "}");
            Put(out, outSize, tmp);
        }

        // A run of doubles, which is what FP, FP12 and the type-O triple all
        // hold once their headers have been read.
        void RenderDoubleGrid(const double* cells, long long rows, long long cols,
                              char* out, int outSize)
        {
            RenderGrid("Double", rows, cols, out, outSize, [&](long long i, char* c, int n)
            {
                double d = 0;
                if (!SafeRead(&d, cells + i, sizeof(d))) return false;
                FormatDouble(d, c, n);
                return true;
            });
        }

        // Copy only the header; the doubles are read one at a time.
        template <class Fp>
        void ReadFp(const void* p, char* out, int outSize)
        {
            Fp h{};
            if (!SafeRead(&h, p, offsetof(Fp, array))) { Put(out, outSize, ""); return; }
            const double* cells = reinterpret_cast<const double*>(
                static_cast<const unsigned char*>(p) + offsetof(Fp, array));
            RenderDoubleGrid(cells, h.rows, h.columns, out, outSize);
        }

        // Handle both the small grid XLOPER and the big grid XLOPER12.
        struct NarrowOper
        {
            using Oper = XLOPER;    using Ref = XLREF;    using MRef = XLMREF;
            static void Str(const void* s, char* out, int n) { ReadAsciiCounted(s, out, n); }
        };
        struct WideOper
        {
            using Oper = XLOPER12;  using Ref = XLREF12;  using MRef = XLMREF12;
            static void Str(const void* s, char* out, int n) { ReadWideCounted(s, out, n); }
        };

        // "R<rw>C<col>:R<rw>C<col>" for one area, 1-based.
        template <class Ref>
        int AppendArea(char* out, int cap, int len, const Ref& a)
        {
            char t[64];
            _snprintf_s(t, _TRUNCATE, "R%lldC%lld:R%lldC%lld",
                        static_cast<long long>(a.rwFirst) + 1, static_cast<long long>(a.colFirst) + 1,
                        static_cast<long long>(a.rwLast) + 1, static_cast<long long>(a.colLast) + 1);
            return Append(out, cap, len, t);
        }

        template <class T>
        void DescribeOper(const void* p, char* out, int outSize, int depth = 0)
        {
            // An array element that is itself an array is not followed further.
            if (depth > 1) { Put(out, outSize, "[...]"); return; }
            typename T::Oper x{};
            if (!SafeRead(&x, p, sizeof(x))) { Put(out, outSize, ""); return; }
            const int t = x.xltype & core::kXlTypeMask;

            switch (t)
            {
            case xltypeNum:     PutDouble(out, outSize, x.val.num);                      return;
            case xltypeStr:     T::Str(x.val.str, out, outSize);                         return;
            case xltypeBool:    Put(out, outSize, x.val.xbool ? "TRUE" : "FALSE");       return;
            case xltypeErr:     Put(out, outSize, ErrName(static_cast<int>(x.val.err))); return;
            case xltypeInt:     PutInt(out, outSize, x.val.w);                           return;
            case xltypeMissing: Put(out, outSize, "Missing");                            return;
            case xltypeNil:     Put(out, outSize, "Empty");                              return;
            case xltypeSRef:
            {
                char tmp[64];
                int len = Append(tmp, sizeof(tmp), 0, "SRef(");
                len = AppendArea(tmp, sizeof(tmp), len, x.val.sref.ref);
                Append(tmp, sizeof(tmp), len, ")");
                Put(out, outSize, tmp);
                return;
            }
            case xltypeRef:
            {
                const void* m = x.val.mref.lpmref;
                using MRef = typename T::MRef;
                using Ref = typename T::Ref;
                MRef head{};
                if (!SafeRead(&head, m, offsetof(MRef, reftbl))) { Put(out, outSize, ""); return; }
                char tmp[kValueMax];
                constexpr int kTail = 48;
                const int body = static_cast<int>(sizeof(tmp)) - kTail;
                int len = Append(tmp, body, 0, "Ref(");
                unsigned shown = 0;
                for (unsigned i = 0; i < head.count; ++i)
                {
                    Ref area{};
                    const void* at = static_cast<const unsigned char*>(m) + offsetof(MRef, reftbl) + i * sizeof(Ref);
                    if (!SafeRead(&area, at, sizeof(area))) break;
                    if (len + 48 >= body) break;
                    if (i) len = Append(tmp, body, len, ",");
                    len = AppendArea(tmp, body, len, area);
                    ++shown;
                }
                if (shown < head.count)
                    len = core::AppendShownMarker(tmp, sizeof(tmp), len, shown, head.count);
                Append(tmp, sizeof(tmp), len, ")");
                Put(out, outSize, tmp);
                return;
            }
            case xltypeMulti:
            {
                const void* arr = x.val.array.lparray;
                const long long rows = x.val.array.rows;
                const long long cols = x.val.array.columns;
                RenderGrid("Variant", rows, cols, out, outSize, [&](long long i, char* c, int n)
                {
                    DescribeOper<T>(static_cast<const unsigned char*>(arr) + i * sizeof(typename T::Oper), c, n, depth + 1);
                    if (!c[0]) Put(c, n, "?");      // an unreadable cell renders as ?
                    return true;
                });
                return;
            }
            default:
                break;
            }
            char tmp[32]; _snprintf_s(tmp, _TRUNCATE, "?xltype%d", t);
            Put(out, outSize, tmp);
        }

        // One value, whatever slot it came from: `bits` is the integer register or stack slot,
        // `dbl` the XMM one. False for a kind with no value of its own. Empty output means the
        // value could not be established.
        bool DescribeValue(Kind k, ULONG64 bits, double dbl, char* out, int outSize)
        {
            const void* p = reinterpret_cast<const void*>(bits);
            switch (k)
            {
            case Kind::Double:    PutDouble(out, outSize, dbl);                                  return true;
            case Kind::Int16:     PutInt(out, outSize, static_cast<short>(bits & 0xFFFF));       return true;
            case Kind::UInt16:    PutInt(out, outSize, static_cast<unsigned short>(bits & 0xFFFF)); return true;
            case Kind::Int32:     PutInt(out, outSize, static_cast<INT32>(bits & 0xFFFFFFFF));   return true;
            case Kind::PtrInt16:  { short  v = 0; if (ReadAt(bits, v)) PutInt(out, outSize, v);    return true; }
            case Kind::PtrInt32:  { INT32  v = 0; if (ReadAt(bits, v)) PutInt(out, outSize, v);    return true; }
            case Kind::PtrDouble: { double v = 0; if (ReadAt(bits, v)) PutDouble(out, outSize, v); return true; }
            case Kind::StrAsciiZ:     ReadAsciiZ(p, out, outSize);                return true;
            case Kind::StrAsciiCount: ReadAsciiCounted(p, out, outSize);          return true;
            case Kind::StrWideZ:      ReadWideZ(p, out, outSize);                 return true;
            case Kind::StrWideCount:  ReadWideCounted(p, out, outSize);           return true;
            case Kind::Fp:            ReadFp<FP>(p, out, outSize);                   return true;
            case Kind::Fp12:          ReadFp<FP12>(p, out, outSize);                 return true;
            case Kind::OperNarrow:    DescribeOper<NarrowOper>(p, out, outSize);     return true;
            case Kind::OperWide:      DescribeOper<WideOper>(p, out, outSize);       return true;
            default:                  return false;
            }
        }
    }

    // ---- one argument ------------------------------------------------------
    void DescribeArg(const Slot& s, const Regs& r, char* out, int outSize)
    {
        if (outSize > 0) out[0] = 0;
        const int pos = s.abiIndex;

        switch (s.kind)
        {
        case Kind::ArrayTriple:
        {
            // Three slots: the head renders the whole thing, the other two say
            // so rather than repeating it.
            if (!s.tripleHead) return;   // the caller renders an O array once, at its first slot
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
            RenderDoubleGrid(reinterpret_cast<const double*>(IntArgAt(r, pos + 2)),
                             rows, cols, out, outSize);
            return;
        }
        case Kind::AsyncHandle:
            Put(out, outSize, "AsyncHandle");
            return;
        default:
            // Position picks the index, type picks the register file (xllregs.h).
            DescribeValue(s.kind, IntArgAt(r, pos), DoubleArgAt(r, pos), out, outSize);
            return;
        }
    }

    // ---- the return value --------------------------------------------------
    void DescribeReturn(Kind k, const Regs& r, char* out, int outSize)
    {
        if (outSize > 0) out[0] = 0;
        switch (k)
        {
        // Nothing to report: an async call answers later, and a void or in-place return has no value.
        case Kind::AsyncHandle: return;
        case Kind::Void:        return;
        default:
            // A Double is in XMM0; rax holds a leftover that would decode as a plausible pointer.
            DescribeValue(k, r.rax, r.xmmRet, out, outSize);
            return;
        }
    }
}
