// XRayXL_GetTraceSummary: what has been traced, as a grid a cell can show.
#include "session.h"
#include "exports.h"
#include "paramparse.h"
#include "xlgrid.h"
#include "core/ascii.h"
#include "core/text.h"
#include "core/render.h"
#include "xll/xllhook.h"
#include "vba/vbatrace.h"
#include "vba/vbapcode.h"
#include "core/tracemodes.h"
#include "emit/csv.h"
#include "xlcall.h"

#include <windows.h>
#include <cstdio>
#include <cwctype>
#include <string>

using namespace app;
using namespace app::params;

namespace
{
    // XRayXL_GetTraceSummary: one row per function called, with its count. Both sources keep the
    // counts on the hot path, so this only formats them. Safe to read while armed: entries are only
    // added and counters only rise.
    //
    // Wildcard, not regex: this runs on every volatile recalc over thousands of procedures, and
    // std::regex allocates and throws on a malformed pattern.
    constexpr int kSumRows = 48;      // data rows before truncation
    constexpr int kSumCols = 4;       // Source, Module, Function, Calls
    // + column header, the status block and one truncation or "nothing" note. The row emitters below never exceed this.
    constexpr int kSumCells = (kSumRows + 8) * kSumCols;
    constexpr int kSumText = 72;

    // The summary's grid -- the same XlGrid, sized for the whole table.
    // SumStr/SumStrA/SumNum stay as the names the body uses.
    __declspec(thread) app::XlGrid<kSumCells, kSumText> t_sum;

    void SumStr (int i, const wchar_t* text) { t_sum.Str(i, text); }
    void SumStrA(int i, const char* text)    { t_sum.StrA(i, text); }
    void SumNum (int i, double v)            { t_sum.Num(i, v); }

    // Case-insensitive wildcard: * any run, ? any one. An empty or missing
    // pattern matches everything.
    bool WildMatch(const wchar_t* pat, const wchar_t* txt)
    {
        const wchar_t* star = nullptr;
        const wchar_t* mark = nullptr;
        while (*txt)
        {
            const wchar_t p = *pat ? towupper(*pat) : 0;
            const wchar_t c = towupper(*txt);
            if (p == L'?' || (p != 0 && p == c)) { ++pat; ++txt; continue; }
            if (p == L'*') { star = pat++; mark = txt; continue; }
            if (star) { pat = star + 1; txt = ++mark; continue; }
            return false;
        }
        while (*pat == L'*') ++pat;
        return *pat == 0;
    }

    // The status footer: what the tool is doing and whether it lost anything. A footer, so a
    // spilled formula still finds the function table at row 2. The counts survive Close(), so a
    // call after Disarm shows what the disarm log recorded. `put` is a template to avoid a
    // std::function allocation.
    template <class Put>
    void AppendSummaryStatusFooter(Put put)
    {
        const bool armed = AnythingArmed();
        put(L"", "state", armed ? "ARMED" : "disarmed", armed ? 1.0 : 0.0);

        const unsigned long long bytes = core::modes::GetBufferBytes();
        char bmode[72];
        if (bytes == 0)
            _snprintf_s(bmode, _TRUNCATE, "synchronous (crash-safe)");
        else
        {
            char size[32];
            core::FormatBytes(bytes, size, sizeof(size));
            _snprintf_s(bmode, _TRUNCATE, "%s ring, BUFFERWHENFULL=%s",
                        size, core::modes::GetPauseOnFull() ? "PAUSE" : "DROP");
        }
        put(L"", "buffer", bmode, static_cast<double>(bytes));

        // THE P-CODE TABLE'S HEALTH, where a cell can read it. It is the one
        // number that says whether the types in this trace can be believed, and
        // until now it existed only in a log line at disarm.
        {
            long long walks = 0, clean = 0;
            vba::PcodeHealth(walks, clean);
            if (walks > 0)
            {
                char h[72];
                _snprintf_s(h, sizeof h, _TRUNCATE, "%lld of %lld procedure(s) walked cleanly",
                            clean, walks);
                put(L"", "p-code", h, walks ? (100.0 * static_cast<double>(clean)
                                                     / static_cast<double>(walks)) : 0.0);
            }
        }

        const long long dropped = emit::csv::RingDrops();
        const long long paused  = emit::csv::RingPauses();
        if (dropped > 0) put(L"", "buffer", "rows dropped (full)", static_cast<double>(dropped));
        if (paused  > 0) put(L"", "buffer", "hot-path pauses (full)", static_cast<double>(paused));

        // The file this session writes. Its name is settled at arm; `rows` says whether anything has reached it.
        const std::wstring wp = emit::csv::Path();
        char fileLeaf[80];
        if (wp.empty()) _snprintf_s(fileLeaf, _TRUNCATE, "(not armed)");
        else
        {
            const std::size_t slash = wp.find_last_of(L'\\');
            const wchar_t* lw = (slash == std::wstring::npos) ? wp.c_str() : wp.c_str() + slash + 1;
            core::NarrowAscii(lw, fileLeaf, static_cast<int>(sizeof fileLeaf) - 1);
        }
        put(L"", "file", fileLeaf, 0);
    }

    LPXLOPER12 GetTraceSummaryBody(LPXLOPER12 filterArg)
    {
        wchar_t pat[64];
        if (IsMissing(filterArg) || ArgType(filterArg) != xltypeStr) { pat[0] = L'*'; pat[1] = 0; }
        else
        {
            const XCHAR* p = filterArg->val.str;
            int n = p ? p[0] : 0;
            if (n > 62) n = 62;
            for (int i = 0; i < n; i++) pat[i] = p[1 + i];
            pat[n] = 0;
            if (n == 0) { pat[0] = L'*'; pat[1] = 0; }
        }

        int row = 0;
        auto put = [&](const wchar_t* src, const char* mod, const char* fn, double calls)
        {
            const int b = row * kSumCols;
            SumStr (b + 0, src);
            SumStrA(b + 1, mod);
            SumStrA(b + 2, fn);
            SumNum (b + 3, calls);
            ++row;
        };

        SumStr(0, L"Source"); SumStr(1, L"Module"); SumStr(2, L"Function"); SumStr(3, L"Calls");
        row = 1;

        int shown = 0, matched = 0;
        // ONE GATE for both sources: the wildcard, the row cap, and the count
        // of what the cap hid.
        auto offer = [&](const wchar_t* src, const char* mod, const char* fn, double calls)
        {
            wchar_t wname[kSumText];
            core::WidenUtf8(fn, wname, kSumText);
            if (!WildMatch(pat, wname)) return;
            ++matched;
            if (shown < kSumRows) { put(src, mod, fn, calls); ++shown; }
        };

        // ---- XLL: every armed target that was actually called ---------------
        for (int i = 0, n = xll::Count(); i < n; i++)
        {
            xll::Target* t = xll::At(i);
            if (t && t->live && t->calls != 0)       // armed but never called is not "traced"
                offer(L"XLL", t->module, t->name, static_cast<double>(t->calls));
        }

        // ---- VBA: every procedure the walk saw ------------------------------
        for (int i = 0, slots = vba::ProcTableSize(); i < slots; i++)
        {
            unsigned long long trailer = 0, calls = 0;
            char mod[384] = {}, fn[256] = {};
            if (!vba::ProcSnapshot(i, trailer, calls, mod, sizeof(mod), fn, sizeof(fn))) continue;
            if (calls == 0) continue;
            // An unresolved name is shown as its trailer address, which at
            // least looks like an address rather than claiming a name.
            char shownName[64];
            if (fn[0]) _snprintf_s(shownName, _TRUNCATE, "%s", fn);
            else       _snprintf_s(shownName, _TRUNCATE, "0x%llX", trailer);
            offer(L"VBA", mod, shownName, static_cast<double>(calls));
        }

        if (shown == 0)
        {
            // Nothing is a real answer, and says so rather than returning an
            // empty grid that reads as a broken formula.
            put(L"", "", "(nothing traced yet)", 0);
        }
        else if (matched > shown)
        {
            // Loud truncation: a silently cut list is a confident wrong answer.
            char more[64];
            _snprintf_s(more, _TRUNCATE, "... %d more -- narrow the filter", matched - shown);
            put(L"", "", more, static_cast<double>(matched));
        }

        // The status footer -- armed state, buffer mode, drops/pauses, file --
        // sized into the grid's spare rows (kSumRows + 6) after any note above.
        AppendSummaryStatusFooter(put);

        return t_sum.AsGrid(row, kSumCols);
    }
}

extern "C" LPXLOPER12 __stdcall XRayXL_GetTraceSummary(LPXLOPER12 filter)
{
    return GuardedOper("XRayXL_GetTraceSummary FAULTED -- contained", GetTraceSummaryBody, filter);
}
