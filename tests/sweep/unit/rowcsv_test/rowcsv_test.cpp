// UNIT TEST for emit::csv::Fragment / emit::csv::kHeader -- the row-to-CSV formatter,
// checked from the INSIDE, with no Excel and no file. The suites' reader-contract
// test checks the same format from the outside (a real trace file); this pins it
// at the source: column count, order, escaping, and the header string.
//
// Built by XRayXL.sln into build\x64\Release\unit\; it needs only rowcsv.cpp.

#include "rowcsv.h"

#include <cstdio>
#include <cstring>
#include <string>
#include <vector>

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what)
    {
        printf("[%s] %s\n", ok ? "PASS" : "FAIL", what);
        if (!ok) ++g_fail;
    }

    // Split a fragment (which ends CRLF) into its comma fields, honouring RFC4180
    // quoting -- the inverse of the escaping, so the test reads the row the way a
    // conformant reader would.
    std::vector<std::string> Fields(const char* frag, int n)
    {
        std::vector<std::string> out;
        std::string cur;
        bool inQ = false;
        int i = 0;
        // strip trailing CRLF
        while (n >= 2 && frag[n - 1] == '\n' && frag[n - 2] == '\r') { n -= 2; break; }
        for (; i < n; ++i)
        {
            const char c = frag[i];
            if (inQ)
            {
                if (c == '"')
                {
                    if (i + 1 < n && frag[i + 1] == '"') { cur += '"'; ++i; }
                    else inQ = false;
                }
                else cur += c;
            }
            else if (c == '"') inQ = true;
            else if (c == ',') { out.push_back(cur); cur.clear(); }
            else cur += c;
        }
        out.push_back(cur);
        return out;
    }
}

int main()
{
    using emit::csv::Row;
    static char buf[emit::csv::kFragMax + 64];

    // ---- the header names the columns the fragment fills, after seq,input ----
    {
        const std::string h = emit::csv::kHeader;
        Check(h.rfind("seq,input,kind,source,", 0) == 0, "header starts seq,input,kind,source");
        Check(h.size() >= 2 && h.substr(h.size() - 2) == "\r\n", "header ends CRLF");
        // header columns = 2 prefixes + 20 fragment fields = 22
        std::string body = h.substr(0, h.size() - 2);
        int commas = 0; for (char c : body) if (c == ',') ++commas;
        Check(commas == 21, "header has 22 columns (21 commas)");
    }

    // ---- a plain row: 20 fields, in order, CRLF-terminated --------------------
    {
        Row r;
        r.kind = "entry"; r.source = "XLL"; r.span = "5"; r.thread = "7"; r.qpc = "99";
        r.parent = "4"; r.depth = "2";
        r.module = "Lib.xll"; r.function = "F"; r.proc = "F"; r.typetext = "QBB";
        r.caller = "cell"; r.callerref = "[B.xlsm]S1!A1"; r.argcount = "2";
        r.args = "1 2"; r.ret = "3"; r.rettype = "Q"; r.outcome = "returned";
        r.ticks = "604"; r.trust = "exit";
        const int n = emit::csv::Fragment(r, buf);
        Check(n >= 2 && buf[n - 2] == '\r' && buf[n - 1] == '\n', "fragment ends CRLF");
        auto f = Fields(buf, n);
        Check(f.size() == 20, "fragment has 20 fields");
        Check(f[0] == "entry" && f[1] == "XLL" && f[8] == "F" && f[18] == "604",
              "fields land in column order");
        // THE COLUMNS THAT MOVED, pinned by POSITION here and by NAME in the
        // header check above: between them, a field that slides one place along
        // cannot pass both.
        Check(f[3] == "4" && f[4] == "2", "parent and depth are 4 and 5, beside span");
        Check(f[11] == "cell" && f[12] == "[B.xlsm]S1!A1", "caller and callerref are 12 and 13");
        Check(f[17] == "returned", "outcome is column 18, after rettype");
        Check(f[18] == "604" && f[19] == "exit", "ticks and trust are the last two columns");
    }

    // ---- escaping: a field with commas/quotes/newlines round-trips -----------
    {
        Row r;
        r.kind = "entry"; r.source = "VBA";
        r.args = "a,b,\"c\",line";           // comma + embedded quotes
        r.trust = "x\r\ny";                    // the LAST column, so a stray newline would also break the row count
        const int n = emit::csv::Fragment(r, buf);
        auto f = Fields(buf, n);
        Check(f.size() == 20, "escaped row still has 20 fields");
        Check(f[14] == "a,b,\"c\",line", "comma+quote field round-trips through the reader");
        // the raw fragment must have quoted the args field
        Check(strstr(buf, "\"a,b,\"\"c\"\",line\"") != nullptr, "args field is quoted with doubled quotes");
        Check(f[19].find('\r') == std::string::npos && f[19].find('\n') == std::string::npos,
              "newlines in a field are neutralised");
    }

    // ---- a newline deep inside a long field is neutralised too ----------------
    {
        // Past the 40th character: a short value passes a writer that only cleans a field's start.
        std::string longArgs(60, 'a');
        longArgs += "\r\nsecond line";
        Row r; r.kind = "entry"; r.source = "VBA"; r.args = longArgs.c_str(); r.trust = "exit";
        const int n = emit::csv::Fragment(r, buf);
        auto f = Fields(buf, n);
        Check(f.size() == 20, "a long field with a newline still has 20 fields");
        Check(f.size() == 20 && f[14].find('\r') == std::string::npos && f[14].find('\n') == std::string::npos,
              "a newline past the 40th character of a field is neutralised");
        int crlf = 0; for (int i = 0; i + 1 < n; ++i) if (buf[i] == '\r' && buf[i + 1] == '\n') ++crlf;
        Check(crlf == 1, "the fragment holds exactly one CRLF, its own terminator");
    }

    // ---- a big field is bounded by kFragMax, not overrun ---------------------
    {
        std::string big(emit::csv::kFragMax * 2, 'x');   // far larger than the buffer
        Row r; r.kind = "entry"; r.source = "XLL"; r.args = big.c_str();
        const int n = emit::csv::Fragment(r, buf);
        Check(n <= emit::csv::kFragMax, "oversized row is truncated within kFragMax, never overruns");
        Check(n >= 2 && buf[n - 2] == '\r' && buf[n - 1] == '\n', "truncated row still ends CRLF");
        // Truncating a value must never drop a column.
        auto fb = Fields(buf, n);
        Check(fb.size() == 20, "an oversized row still has 20 fields");
        Check(fb.size() == 20 && fb[0] == "entry" && fb[1] == "XLL",
              "an oversized row keeps its leading columns");
    }

    // ---- a huge value in the LAST column -------------------------------------
    {
        std::string big(emit::csv::kFragMax * 2, 'y');
        Row r; r.kind = "exit"; r.source = "VBA"; r.trust = big.c_str();
        const int n = emit::csv::Fragment(r, buf);
        Check(n <= emit::csv::kFragMax, "a huge last column stays within kFragMax");
        auto f = Fields(buf, n);
        Check(f.size() == 20, "a huge last column still leaves 20 fields");
    }

    // ---- a huge value mid-row must not cost the columns after it -------------
    {
        std::string big(emit::csv::kFragMax * 2, 'z');
        Row r; r.kind = "exit"; r.source = "VBA"; r.args = big.c_str();
        r.ret = "42"; r.rettype = "Long"; r.outcome = "returned"; r.ticks = "604"; r.trust = "exit";
        const int n = emit::csv::Fragment(r, buf);
        auto f = Fields(buf, n);
        Check(f.size() == 20 && f[15] == "42" && f[16] == "Long" && f[17] == "returned" &&
              f[18] == "604" && f[19] == "exit",
              "a huge args value keeps ret, rettype, outcome, ticks and trust");
        Check(f.size() == 20 && f[14].size() >= 3 && f[14].compare(f[14].size() - 3, 3, "...") == 0,
              "the cut args value ends with ... to say so");
    }

    // ---- empty row: all fields empty, still 20 of them -----------------------
    {
        Row r;
        const int n = emit::csv::Fragment(r, buf);
        auto f = Fields(buf, n);
        Check(f.size() == 20, "an all-empty row is 20 empty fields");
        bool allEmpty = true; for (auto& s : f) if (!s.empty()) allEmpty = false;
        Check(allEmpty, "empty row's fields are all empty");
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
