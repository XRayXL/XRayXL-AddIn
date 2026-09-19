// UNIT TEST for emit::csv::Fragment / FragmentSize / kHeader -- the row-to-CSV formatter,
// checked from the INSIDE, with no Excel and no file. The suites' reader-contract
// test checks the same format from the outside (a real trace file); this pins it
// at the source: column count, order, escaping, and the header string.
//
// Built by XRayXL.sln into build\x64\Release\unit\; it needs only rowcsv.cpp.

#include "rowcsv.h"
#include "rowjson.h"

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

namespace
{
    // The fragment for `r`, in a buffer sized by FragmentSize, which must be exact.
    std::vector<char> g_buf;
    int Frag(const emit::csv::Row& r)
    {
        const std::size_t size = emit::csv::FragmentSize(r);
        g_buf.assign(size + 1, 0);
        const std::size_t n = emit::csv::Fragment(r, g_buf.data());
        if (n != size) Check(false, "FragmentSize is exactly what Fragment writes");
        return static_cast<int>(n);
    }
}

int main()
{
    using emit::csv::Row;

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
        const int n = Frag(r); const char* buf = g_buf.data();
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
        const int n = Frag(r); const char* buf = g_buf.data();
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
        const int n = Frag(r); const char* buf = g_buf.data();
        auto f = Fields(buf, n);
        Check(f.size() == 20, "a long field with a newline still has 20 fields");
        Check(f.size() == 20 && f[14].find('\r') == std::string::npos && f[14].find('\n') == std::string::npos,
              "a newline past the 40th character of a field is neutralised");
        int crlf = 0; for (int i = 0; i + 1 < n; ++i) if (buf[i] == '\r' && buf[i + 1] == '\n') ++crlf;
        Check(crlf == 1, "the fragment holds exactly one CRLF, its own terminator");
    }

    // ---- a big field is written whole, never cut -----------------------------
    {
        std::string big(1 << 20, 'x');
        Row r; r.kind = "entry"; r.source = "XLL"; r.args = big.c_str();
        r.ret = "42"; r.rettype = "Long"; r.outcome = "returned"; r.ticks = "604"; r.trust = "exit";
        const int n = Frag(r); const char* buf = g_buf.data();
        auto f = Fields(buf, n);
        Check(f.size() == 20 && f[14] == big, "a 1 MB field is written whole");
        Check(f.size() == 20 && f[15] == "42" && f[16] == "Long" && f[17] == "returned" &&
              f[18] == "604" && f[19] == "exit",
              "a big args value keeps ret, rettype, outcome, ticks and trust");
    }

    // ---- escaping is counted: a field of quotes doubles and still fits exactly
    {
        std::string quotes(1000, '"');
        Row r; r.kind = "exit"; r.source = "VBA"; r.trust = quotes.c_str();
        const int n = Frag(r); const char* buf = g_buf.data();
        auto f = Fields(buf, n);
        Check(f.size() == 20 && f[19] == quotes, "a last column of quotes round-trips whole");
    }

    // ---- JSON Lines: the same row as one object, after the writer's prefixes ----
    {
        Row r;
        r.kind = "exit"; r.source = "VBA"; r.span = "5"; r.parent = "0"; r.depth = "1";
        r.thread = "7"; r.qpc = "99"; r.module = "Mod\"1"; r.function = "F";
        r.ret = "{\"t\":\"Long\",\"v\":5}"; r.rettype = "Long"; r.outcome = "returned";
        r.ticks = "604"; r.trust = "exit";
        core::TextBuf out;
        out.Append("{\"seq\":1,\"input\":1,");
        emit::json::AppendRow(r, out);
        printf("JSON %.*s\n", static_cast<int>(out.Len() - 2), out.Text());
        Check(std::string(out.Text()) ==
              "{\"seq\":1,\"input\":1,\"kind\":\"exit\",\"source\":\"VBA\",\"span\":5,\"parent\":0,"
              "\"depth\":1,\"thread\":7,\"qpc\":99,\"module\":\"Mod\\\"1\",\"function\":\"F\","
              "\"ret\":{\"t\":\"Long\",\"v\":5},\"rettype\":\"Long\",\"outcome\":\"returned\","
              "\"ticks\":604,\"trust\":\"exit\"}\r\n",
              "a JSON row: integers as numbers, ret as JSON, empty fields omitted, one line");
        out.Release();
    }

    // ---- empty row: all fields empty, still 20 of them -----------------------
    {
        Row r;
        const int n = Frag(r); const char* buf = g_buf.data();
        auto f = Fields(buf, n);
        Check(f.size() == 20, "an all-empty row is 20 empty fields");
        bool allEmpty = true; for (auto& s : f) if (!s.empty()) allEmpty = false;
        Check(allEmpty, "empty row's fields are all empty");
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
