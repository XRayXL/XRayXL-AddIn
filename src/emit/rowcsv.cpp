#include "rowcsv.h"
#include <cstring>

namespace emit
{

namespace csv
{
    // The header and the column order together. docs/TraceRowModel.md quotes the header byte
    // for byte and the suites' reader refuses a file that differs.
    //
    // `seq` is the writer's, dense in file order; `input` is the producer's, so its holes show
    // where rows were dropped. Both are prefixes added outside Fragment, so kColumns counts
    // only kind..trust.
    const char* const kHeader =
        "seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,"
        "caller,callerref,argcount,args,ret,rettype,outcome,ticks,tracerticks,trust\r\n";
    const char* const kHeaderBreaks =
        "seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,"
        "caller,callerref,argcount,args,ret,rettype,outcome,ticks,tracerticks,trust,breaks\r\n";

    namespace
    {
        constexpr int kColumns = 22;    // fragment fields, after the seq/input prefixes; the last is optional
        void InOrder(const Row& r, const char* (&f)[kColumns])
        {
            f[0]  = r.kind;     f[1]  = r.source;   f[2]  = r.span;     f[3]  = r.parent;
            f[4]  = r.depth;    f[5]  = r.thread;   f[6]  = r.qpc;      f[7]  = r.module;
            f[8]  = r.function; f[9]  = r.proc;     f[10] = r.typetext; f[11] = r.caller;
            f[12] = r.callerref; f[13] = r.argcount; f[14] = r.args;    f[15] = r.ret;
            f[16] = r.rettype;  f[17] = r.outcome;  f[18] = r.ticks;
            f[19] = r.tracerticks; f[20] = r.trust;  f[21] = r.breaks;
        }

        bool NeedsQuote(const char* src)
        {
            for (const char* p = src; *p; p++)
                if (*p == ',' || *p == '"' || *p == '\n' || *p == '\r') return true;
            return false;
        }

        std::size_t EscapedSize(const char* src)
        {
            std::size_t n = NeedsQuote(src) ? 2 : 0;
            for (const char* p = src; *p; p++) n += (*p == '"') ? 2 : 1;
            return n;
        }

        // Minimal RFC4180 escaping. Values come out of an add-in's memory and
        // can contain anything, so this is not optional. `dst` holds EscapedSize(src).
        std::size_t Escape(const char* src, char* dst)
        {
            const bool quote = NeedsQuote(src);
            std::size_t n = 0;
            if (quote) dst[n++] = '"';
            for (const char* p = src; *p; p++)
            {
                if (*p == '"')                     { dst[n++] = '"'; dst[n++] = '"'; }
                else if (*p == '\r' || *p == '\n') dst[n++] = ' ';
                else                               dst[n++] = *p;
            }
            if (quote) dst[n++] = '"';
            return n;
        }
    }

    std::size_t FragmentSize(const Row& row, bool breaks)
    {
        const char* fields[kColumns];
        InOrder(row, fields);
        const int cols = breaks ? kColumns : kColumns - 1;
        std::size_t n = cols - 1 + 2;         // the commas, then CRLF
        for (int i = 0; i < cols; i++) n += EscapedSize(fields[i] ? fields[i] : "");
        return n;
    }

    std::size_t Fragment(const Row& row, char* out, bool breaks)
    {
        const char* fields[kColumns];
        InOrder(row, fields);
        const int cols = breaks ? kColumns : kColumns - 1;
        std::size_t n = 0;
        // Always emit every column: the reader rejects a short row.
        for (int i = 0; i < cols; i++)
        {
            if (i) out[n++] = ',';
            n += Escape(fields[i] ? fields[i] : "", out + n);
        }
        out[n++] = '\r';
        out[n++] = '\n';
        out[n] = 0;
        return n;
    }
}
}   // namespace emit
