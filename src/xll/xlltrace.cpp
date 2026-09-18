#include "xlltrace.h"
#include "core/caller.h"
#include "xllhook.h"
#include "xllregs.h"
#include "xlldecode.h"
#include "xllregistry.h"
#include "emit/csv.h"
#include "core/excel_api.h"
#include "core/log.h"
#include "core/crashlog.h"
#include "xlcall.h"
#include "core/clock.h"

#include <cstdio>

namespace xll
{
    // ---- WHAT THE HOT PATH DECODES, latched at arm --------------------------
    // Read ONCE from core::modes::, so the hot path never sees a value change under
    // it. A disabled decode leaves its column EMPTY and is COUNTED, keeping "we
    // did not look" distinguishable from "we looked and found nothing".
    volatile LONG g_capArgs = 1;
    volatile LONG g_capRet  = 1;
    // DEPTH=TOP: only the OUTERMOST call on a thread emits a row, the one Excel
    // itself initiated. An add-in that re-enters Excel through xlUDF nests
    // genuinely (TxCallsBack2 reaches depth 3), so TOP is how a user says they
    // want only the call the sheet made.
    volatile LONG g_topOnly = 0;

    void SetCapture(bool args, bool retval)
    {
        InterlockedExchange(&g_capArgs, args   ? 1 : 0);
        InterlockedExchange(&g_capRet,  retval ? 1 : 0);
    }

    void SetTopOnly(bool topOnly)
    {
        InterlockedExchange(&g_topOnly, topOnly ? 1 : 0);
    }

    namespace
    {
        // An entry row with no exit is indistinguishable in the file from a
        // call that hung, so Disarm reports this when non-zero.
        volatile LONG64 g_exitsDropped = 0;
        // Rows lost to a fault while decoding, and frames closed because an
        // exception unwound past the thunk without an exit.
        volatile LONG64 g_recorderFaults = 0;
        volatile LONG64 g_framesResynced = 0;

        // Per-thread state. TLS only: multithreaded calculation puts several threads inside one
        // hooked function at once.
        struct Frame
        {
            unsigned long long span = 0;
            long long          startQpc = 0;
            bool               recorded = false;   // did its ENTRY get written?
            ULONG_PTR          sp = 0;             // the thunk frame, for resynchronising
        };

        struct ThreadState
        {
            int    depth = 0;                 // call nesting on this thread
            bool   inside = false;            // we are inside our own recorder
            static const int kMaxDepth = 64;
            Frame  frames[kMaxDepth];
        };
        __declspec(thread) ThreadState t_state;

        long long Qpc()
        {
            return core::QpcNow();
        }


        // The three columns every row stamps the same way.
        struct Stamp
        {
            char span[24], tid[16], qpc[24];
            Stamp(unsigned long long s, long long q)
            {
                _snprintf_s(span, _TRUNCATE, "%llu", s);
                _snprintf_s(tid,  _TRUNCATE, "%lu", GetCurrentThreadId());
                _snprintf_s(qpc,  _TRUNCATE, "%lld", q);
            }
        };

        // ---- record writing ---------------------------------------------------
        void WriteEntry(Target* t, const Regs& r, unsigned long long span, unsigned long long parent,
                        int depth, long long qpc)
        {
            const Stamp st(span, qpc);
            char parentb[24], depthb[16];
            _snprintf_s(parentb, _TRUNCATE, "%llu", parent);
            _snprintf_s(depthb, _TRUNCATE, "%d", depth);
            char argcbuf[8];
            _snprintf_s(argcbuf, _TRUNCATE, "%d", t->plan.paramCount);

            // Decoded in one place for both sources (core/caller.h). Asking costs a round trip
            // into Excel once per entry, inside the span being timed.
            core::Caller who;      // NOT {} -- ReadCaller fills it
            core::ReadCaller(who);

            // One field, "a<slot>:<code>=<value>" joined by spaces -- the grammar VBA's args use -- so
            // the CSV stays rectangular whatever the arity. An O array renders once, at its first slot.
            char args[2048]; args[0] = 0;
            int n = 0;
            const bool wantArgs = (InterlockedCompareExchange(&g_capArgs, 0, 0) != 0);
            int described = 0;
            for (int i = 0; wantArgs && i < t->plan.describedCount && n < static_cast<int>(sizeof(args)) - 160; i++)
            {
                const Slot& s = t->plan.slots[i];
                if (s.kind == Kind::ArrayTriple && !s.tripleHead) { described = i + 1; continue; }
                char v[kValueRenderMax] = { 0 };
                DescribeArg(s, r, v, sizeof(v));
                // The `?` is unreachable: a plan whose codes did not all parse is never hooked.
                const int w = _snprintf_s(args + n, sizeof(args) - n, _TRUNCATE, "%sa%d:%s=%s",
                                          n ? " " : "", i + 1, s.code[0] ? s.code : "?", v);
                if (w < 0) break;
                n += w;
                described = i + 1;
            }
            // Say when the list stops short of the real arity.
            if (wantArgs && described < t->plan.slotCount && n < static_cast<int>(sizeof(args)) - 64)
            {
                _snprintf_s(args + n, sizeof(args) - n, _TRUNCATE,
                            "%s...(%d of %d slots described)", n ? " " : "", described, t->plan.slotCount);
            }

            // No seq: emit::csv::WriteRow allocates it. argcount is the arity, not the number
            // decoded, so ARGS=FALSE does not make a call look nullary.
            emit::csv::Row row;
            row.kind = "entry";  row.source = "XLL";  row.span = st.span;  row.thread = st.tid;  row.qpc = st.qpc;
            row.parent = parentb;  row.depth = depthb;
            row.module = t->module;  row.function = t->name;
            row.proc = t->procName;  row.typetext = t->plan.signature;
            row.caller = who.kind;  row.callerref = who.desc;
            row.argcount = argcbuf;  row.args = args;
            emit::csv::WriteRow(row);
        }

        void WriteExit(Target* t, const Regs& r, unsigned long long span, unsigned long long parent,
                       int depth, long long qpc, long long startQpc)
        {
            const Stamp st(span, qpc);
            char parentb[24], depthb[16];
            _snprintf_s(parentb, _TRUNCATE, "%llu", parent);
            _snprintf_s(depthb, _TRUNCATE, "%d", depth);
            char durbuf[32];

            // An async function has not produced its answer yet, so the span measures only the
            // dispatch. `ticks` is left empty and `trust` says `async`.
            const char* trustText;
            if (t->plan.async) { durbuf[0] = 0; trustText = "async"; }
            else
            {
                _snprintf_s(durbuf, _TRUNCATE, "%lld", qpc - startQpc);
                // The detour fires on the RETURN PATH, so this row existing is
                // the evidence the call came back -- the same fact VBA's exit
                // opcode gives, and it earns the same word.
                trustText = "exit";
            }

            char ret[kValueRenderMax] = { 0 };
            const bool wantRet = (InterlockedCompareExchange(&g_capRet, 0, 0) != 0);
            if (wantRet) DescribeReturn(t->plan.returnKind, r, ret, sizeof(ret));

            // No caller on an exit row: it is the same activation as the entry
            // row it pairs with, which already named one.
            emit::csv::Row row;
            row.kind = "exit";  row.source = "XLL";  row.span = st.span;  row.thread = st.tid;  row.qpc = st.qpc;
            row.parent = parentb;  row.depth = depthb;
            row.module = t->module;  row.function = t->name;
            row.proc = t->procName;
            row.ret = ret;
            // The registered return code with its flags ("Q$"); empty only when the return was not
            // captured, since an unparsed plan is never hooked.
            row.rettype = wantRet ? t->plan.returnText : "";
            // Only the return path writes this row, so `returned` is always true.
            row.outcome = "returned";
            row.ticks = durbuf;  row.trust = trustText;
            emit::csv::WriteRow(row);
        }
    }

    // The recorders. Both SEH-wrapped: they decode pointers we did not create on somebody
    // else's calc thread, so a fault must cost the row and nothing more.

    extern "C" void XRayOnEntry(void* target, void* regs)
    {
        Target* t = static_cast<Target*>(target);
        const Regs* r = static_cast<const Regs*>(regs);

        // Depth is always balanced: entry increments and exit decrements whether or not the row
        // is written. A frame whose thunk sits at or below this one on the stack was unwound
        // without an exit, so close it before going deeper.
        const ULONG_PTR sp = reinterpret_cast<ULONG_PTR>(regs);
        while (t_state.depth > 0)
        {
            // Frames past the table were never recorded; they are stale if the deepest recorded one is.
            const int top = t_state.depth > ThreadState::kMaxDepth ? ThreadState::kMaxDepth : t_state.depth;
            if (t_state.frames[top - 1].sp > sp) break;
            t_state.depth = top - 1;
            InterlockedIncrement64(&g_framesResynced);
        }

        const int d = t_state.depth;
        t_state.depth++;
        if (d >= ThreadState::kMaxDepth) return;

        Frame& f = t_state.frames[d];
        f.recorded = false;
        f.span = 0;          // a frame that is not recorded must never lend a stale span as a parent
        f.sp = sp;

        // Reentrancy: our own decoding calls back into Excel, which can reach
        // a hooked function. Without this the first traced call recurses.
        if (t_state.inside) return;

        t_state.inside = true;
        __try
        {
            // Counted before the TOP filter; capped and re-entrant entries returned above.
            InterlockedIncrement64(&t->calls);

            // DEPTH=TOP drops every nested call. The exit is not a separate
            // decision: it writes only when f.recorded says its entry did, so
            // an entry suppressed here can never leave a dangling exit.
            if (d == 0 || InterlockedCompareExchange(&g_topOnly, 0, 0) == 0)
            {
                f.span = emit::csv::NextSpan();
                f.startQpc = Qpc();
                // depth counts this thread's XLL frames; parent is the frame below, 0 at the top.
                WriteEntry(t, *r, f.span, d ? t_state.frames[d - 1].span : 0, d + 1, f.startQpc);
                f.recorded = true;
            }
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            InterlockedIncrement64(&g_recorderFaults);
        }
        t_state.inside = false;
    }

    extern "C" void XRayOnExit(void* target, void* regs)
    {
        Target* t = static_cast<Target*>(target);
        const Regs* r = static_cast<const Regs*>(regs);

        if (t_state.depth <= 0) return;          // never seen an entry: refuse
        t_state.depth--;
        const int d = t_state.depth;
        if (d >= ThreadState::kMaxDepth) return;

        Frame& f = t_state.frames[d];

        // An exit is written only if its entry was emitted, so a filtered entry
        // cannot leave an exit that invents a span. An entry the ring later drops
        // WAS emitted, so its exit is still written; `input` locates the loss.
        if (!f.recorded) return;

        if (t_state.inside) { InterlockedIncrement64(&g_exitsDropped); return; }

        t_state.inside = true;
        __try
        {
            WriteExit(t, *r, f.span, d ? t_state.frames[d - 1].span : 0, d + 1, Qpc(), f.startQpc);
        }
        __except (EXCEPTION_EXECUTE_HANDLER)
        {
            InterlockedIncrement64(&g_recorderFaults);
        }
        t_state.inside = false;
        f.recorded = false;
    }

    long long ExitsDropped() { return InterlockedCompareExchange64(&g_exitsDropped, 0, 0); }
    long long RecorderFaults() { return InterlockedCompareExchange64(&g_recorderFaults, 0, 0); }
    long long FramesResynced() { return InterlockedCompareExchange64(&g_framesResynced, 0, 0); }
}
