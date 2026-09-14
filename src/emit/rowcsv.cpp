#include "rowcsv.h"
#include <cstring>

namespace emit
{

namespace csv
{
    // THE HEADER AND THE COLUMN ORDER, TOGETHER, so changing one puts the other
    // on the next line. docs/TraceRowModel.md quotes the header byte for byte
    // and the suites' reader refuses a file that differs.
    //
    // TWO SEQUENCE NUMBERS: `seq` is the WRITER's -- dense and strictly
    // increasing in file order, the reliable one -- and `input` is the
    // PRODUCER's, so a dropped row consumes a value that never lands and its
    // HOLES show where data was lost. Both are prefixes added outside Fragment,
    // so kColumns below counts only kind..trust.
    const char* const kHeader =
        "seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,"
        "caller,callerref,argcount,args,ret,rettype,outcome,ticks,trust\r\n";

    namespace
    {
        constexpr int kColumns = 20;    // fragment fields, after the seq/input prefixes
        constexpr int kColumnFloor = 512;   // bytes every later column keeps
        void InOrder(const Row& r, const char* (&f)[kColumns])
        {
            f[0]  = r.kind;     f[1]  = r.source;   f[2]  = r.span;     f[3]  = r.parent;
            f[4]  = r.depth;    f[5]  = r.thread;   f[6]  = r.qpc;      f[7]  = r.module;
            f[8]  = r.function; f[9]  = r.proc;     f[10] = r.typetext; f[11] = r.caller;
            f[12] = r.callerref; f[13] = r.argcount; f[14] = r.args;    f[15] = r.ret;
            f[16] = r.rettype;  f[17] = r.outcome;  f[18] = r.ticks;
            f[19] = r.trust;
        }

        // Minimal RFC4180 escaping. Values come out of an add-in's memory and
        // can contain anything, so this is not optional. A value that does not
        // fit ends with "...", so a cut reads as one.
        int Escape(const char* src, char* dst, int dstSize)
        {
            bool needQuote = false;
            for (const char* p = src; *p; p++)
                if (*p == ',' || *p == '"' || *p == '\n' || *p == '\r') { needQuote = true; break; }

            // What the value itself may use: less the NUL, the quotes, and a cut's "...".
            const int body = dstSize - 1 - (needQuote ? 2 : 0) - 3;
            if (body < 0) { if (dstSize > 0) dst[0] = 0; return 0; }

            int n = 0, used = 0;
            if (needQuote) dst[n++] = '"';
            const char* p = src;
            for (; *p; p++)
            {
                const int w = (*p == '"') ? 2 : 1;
                if (used + w > body) break;
                if (*p == '"')                       { dst[n++] = '"'; dst[n++] = '"'; }
                else if (*p == '\r' || *p == '\n') dst[n++] = ' ';
                else                                 dst[n++] = *p;
                used += w;
            }
            if (*p) { dst[n++] = '.'; dst[n++] = '.'; dst[n++] = '.'; }
            if (needQuote) dst[n++] = '"';
            dst[n] = 0;
            return n;
        }
    }

    int Fragment(const Row& row, char* out)
    {
        const char* fields[kColumns];
        InOrder(row, fields);
        int n = 0;
        // Always emit every column: the reader rejects a short row.
        for (int i = 0; i < kColumns; i++)
        {
            if (i) out[n++] = ',';
            // Every later column keeps a comma and kColumnFloor bytes, so one huge
            // value cannot empty the rest; then CRLF and the NUL.
            const int reserve = (kColumns - i - 1) * (1 + kColumnFloor) + 4;
            const int room    = kFragMax - n - reserve;
            if (room > 1)
                n += Escape(fields[i] ? fields[i] : "", out + n, room);
        }
        out[n++] = '\r';
        out[n++] = '\n';
        return n;
    }
}
}   // namespace emit
