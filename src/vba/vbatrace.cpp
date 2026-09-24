#include "vbatrace.h"
#include "vbatrace_internal.h"

#include "emit/csv.h"
#include "vbaidentity.h"
#include "vbaargs.h"
#include "vbaretdecode.h"
#include "vbapcode.h"
#include "vbaboundary.h"
#include "core/caller.h"
#include "vbaregs.h"
#include "vbatrailer.h"
#include "vbaproctable.h"
#include "core/clock.h"
#include "core/hash.h"
#include "core/safemem.h"
#include "core/valueformat.h"
#include "core/tlsstack.h"

#include <windows.h>
#include <cstring>
#include <cstdio>
#include <malloc.h>   // _resetstkoflw
#include <intrin.h>   // _AddressOfReturnAddress

namespace vba
{
    inline void Bump(std::uint64_t& counter)
    {
        InterlockedIncrement64(reinterpret_cast<volatile LONG64*>(&counter));
    }

    inline void Bump(volatile std::uint64_t& counter)
    {
        InterlockedIncrement64(reinterpret_cast<volatile LONG64*>(&counter));
    }

    namespace
    {
        // Exit opcodes with no return-kind mapping, kept by value so the gap can be named.
        SeenTable g_unmappedExitOps;

        // (store opcode, operand offset) seen on untyped returns, so the mapping grows from
        // evidence rather than a guess.
        SeenTable g_retStores;

        // Every exit opcode, mapped or not: ProcedureEnd uses a deny-list, so an unknown
        // mid-procedure exit would close wrongly (vbaboundary.h).
        SeenTable g_exitOpsSeen;

        const std::uint8_t kOutReturned = 0;
        const std::uint8_t kOutThrew    = 1;
        const std::uint8_t kOutUnwound  = 2;
        const std::uint8_t kOutHandled  = 3;
        // `End` kills the session with no epilogue anywhere: these frames neither returned
        // nor threw, and `returned` would be a confident wrong cell.
        const std::uint8_t kOutAbandoned = 4;
        // An unhandled error left VBA through a frame Excel called for a cell.
        const std::uint8_t kOutUnhandledErr = 5;


        const char* OutcomeName(std::uint8_t o)
        {
            switch (o)
            {
            case kOutThrew:   return "threw";
            case kOutUnwound: return "unwound";
            case kOutHandled: return "handled";
            case kOutAbandoned: return "abandoned";
            case kOutUnhandledErr: return "unhandled";
            default:          return "returned";
            }
        }

        struct Frame
        {
            std::uint64_t trailer;
            std::uint64_t startTicks;
            std::uint64_t tracerAtEntry;   // core::TracerTicks() when the entry was stamped
            std::uint64_t stmtsAtEntry;
            std::uint64_t sp;            // interpreter rsp when this frame was seen
            std::uint64_t span;          // pairs this frame's entry and exit rows
            std::uint64_t parent;        // the span we were called from, 0 at the top
            // The depth rows report: 1 for a chain Excel started inside another's DoEvents,
            // though that chain still sits above it on the shadow stack.
            int           shownDepth;
            // How this activation ended:
            //   returned  it finished normally
            //   threw     an error left it, and it was the innermost frame to go
            //   unwound   the error passed through it; it ran nothing after it
            //   handled   it ran again after the error, so it caught it
            std::uint8_t  outcome;
            // How often this activation stopped at a breakpoint in the editor.
            std::uint32_t breaks;

            // Excel started this frame for a worksheet function, so an error leaving it
            // becomes a cell's #VALUE! rather than a VBA caller's error.
            bool          fromExcel;
            // The cell xlfCaller named for this frame, hashed; 0 when not a cell.
            // A frame whose cell differs from the one beneath is a fresh Excel entry.
            std::uint64_t callerHash;

            // Only a ByRef parameter makes the exit re-read worth its cost.
            bool          hasByRef;

            // A hash, not the text: the text is up to 64 KB and the shadow stack is fixed-size
            // TLS. The "before" text is already on the entry row.
            std::uint64_t argsHash;
            // A ByRef argument was written as its shape alone at entry, so no exit can compare.
            bool          argsShapeOnly;

            // No caller gate: the exit opcode says whether a result exists and of what kind.
        };

        // No member initialisers, so it can be TLS with no dynamic initialiser. `generation`
        // starts at 0 and g_generation at 1, so State() initialises the block on first use.
        struct ThreadState
        {
            core::TlsStack<Frame, 256> stack;
            int    depth;
            std::uint64_t current;      // trailer of the running procedure
            std::uint64_t currentSp;    // and the rsp we last saw it at
            std::uint64_t statements;   // this thread's statement count
            std::uint32_t generation;   // which arming session this belongs to

            // An error in flight, set when the thrower closes. A frame predating the throw that
            // runs a statement again caught it; one that ran nothing since unwound.
            bool          errActive;
            // DoEvents open on this thread, and the frame depth each call was made at, so a
            // procedure Excel runs while a frame waits there can be told from one it called.
            int           doEventsMarks;
            int           doEventsAt[8];

            // The thrower's span. Spans only increase, so a lower span predates the throw and
            // can be the handler; a higher one opened during the unwind.
            std::uint64_t errSpan;
        };

        // Hooks run on many threads at once. The storage itself is thread-local, not a pointer.
        __declspec(thread) ThreadState t_state;

        // The argument and return columns' buffers: per thread, on the heap, so ArgCapture
        // stays tiny on the deeply nested interpreter stack. Leaked at thread exit.
        __declspec(thread) core::TextBuf t_argRender;
        __declspec(thread) core::TextBuf t_retRender;

        ProcTable g_procs;
        Totals    g_totals;

        // Bumped by ResetTracing: a thread cannot clear another's TLS, so each clears its own
        // when it sees a stale generation.
        volatile LONG g_generation = 1;

        // Rows for depth-1 frames only. The totals still count every frame.
        volatile LONG g_topLevelOnly = 0;
        // Argument capture follows pointers and walks bytecode, the largest fault surface, so it
        // can be turned off; a disabled decode is counted, never silently blank.
        volatile LONG g_capArgs = 1;

        // ByRef is VBA's default, so a procedure filling in a parameter is ordinary code. The
        // type table marks it with "&".
        bool SignatureHasByRef(const char* sig)
        {
            for (const char* c = sig; *c; ++c) if (*c == '&') return true;
            return false;
        }
        volatile LONG g_capRet  = 1;


        struct ResolvedName
        {
            // Matches Proc's buffers above; this one is a stack temporary.
            char qualModule[384];
            char function[256];
        };

        // The shared name, copied while no publisher holds it. Empty if one does.
        void SnapshotName(Proc* p, ResolvedName& out)
        {
            out.qualModule[0] = 0; out.function[0] = 0;
            if (!p) return;
            for (int spin = 0; spin < 64; ++spin)
            {
                if (InterlockedCompareExchange(&p->publishing, 1, 0) == 0)
                {
                    strncpy_s(out.qualModule, p->qualModule, _TRUNCATE);
                    strncpy_s(out.function,   p->function,   _TRUNCATE);
                    InterlockedExchange(&p->publishing, 0);
                    return;
                }
                YieldProcessor();
            }
        }

        // One sequence across VBA and XLL rows: `qpc` stamps when the event happened, `seq`
        // when the row is written. Runs on a frame change, never per statement.
        void EmitRow(const char* kind, const Frame& f, std::uint64_t qpc,
                     std::uint64_t durationTicks, int depth, Proc* p,
                     const ArgCapture* args = nullptr,
                     const ResolvedName* nm = nullptr,
                     const core::Caller* who = nullptr,
                     const char* retText = "", const char* retType = "",
                     const char* closedBy = "",
                     // Exit rows carry values only: `argcount` and `typetext` belong to
                     // the signature, already on the paired entry row.
                     bool argValuesOnly = false)
        {
            // The only part that touches a file; with it closed the hooks and totals still
            // run, which rules the emitter in or out of a crash.
            if (!emit::csv::IsOpen()) return;
            char spanb[24], tidb[16], qpcb[24], trailerb[32], ticksb[24], tracerb[24] = {};
            char parentb[24], depthb[16];
            _snprintf_s(spanb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(f.span));
            _snprintf_s(tidb, _TRUNCATE, "%lu", GetCurrentThreadId());
            _snprintf_s(qpcb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(qpc));
            _snprintf_s(trailerb, _TRUNCATE, "0x%llX",
                        static_cast<unsigned long long>(f.trailer));
            // `parent` is a span, so a missing row shows as a dangling reference; rebuilding
            // from depth and row order would attach everything after a gap to the wrong caller.
            _snprintf_s(parentb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(f.parent));
            _snprintf_s(depthb, _TRUNCATE, "%d", depth);
            // Exit row only. `trust` decides whether `ticks` is a measurement or an upper bound:
            // only the exit opcode fires as an activation ends; the backstop fires at the next
            // statement and the flush at disarm.
            const char* trustText = "";
            if (durationTicks)
            {
                _snprintf_s(ticksb, _TRUNCATE, "%llu",
                            static_cast<unsigned long long>(durationTicks));
                trustText = (closedBy && closedBy[0]) ? closedBy : "backstop";
                // Read at the exit's stamp: this hook's own time is added only when it returns.
                _snprintf_s(tracerb, _TRUNCATE, "%llu",
                            static_cast<unsigned long long>(core::TracerTicks() - f.tracerAtEntry));
            }
            else
                ticksb[0] = 0;

            // Arguments and the declared signature, when they checked out.
            char argcb[16] = {};
            const char* argsText = "";
            const char* sigText  = "";
            if (args && args->ok)
            {
                argsText = args->text->Text();
                if (!argValuesOnly)
                {
                    _snprintf_s(argcb, _TRUNCATE, "%d", args->params);
                    sigText = args->signature;   // "(Long,String)", or empty
                }
            }

            // `proc` is always the trailer, telling same-named procedures apart. The caller's
            // stack-local name is preferred over the shared Proc entry, which every push rewrites.
            ResolvedName shared{};
            if (p && !(nm && nm->qualModule[0] && nm->function[0])) SnapshotName(p, shared);
            const char* mod  = (nm && nm->qualModule[0]) ? nm->qualModule : shared.qualModule;
            const char* func = (nm && nm->function[0])   ? nm->function
                             : (shared.function[0] ? shared.function : trailerb);

            // Only an entry row carries this. `cell` and `sheet` stay empty for anything that
            // is not a real cell on a real sheet (core/caller.h).
            const char* whoKind = who ? who->kind : "";
            const char* whoRef  = who ? who->desc : "";

            // No seq: emit::csv::WriteRow allocates it under the lock that orders
            // the file, so both sources share one sequence.
            emit::csv::Row row;
            row.kind = kind;  row.source = "VBA";  row.span = spanb;  row.thread = tidb;  row.qpc = qpcb;
            row.module = mod;  row.function = func;  row.proc = trailerb;  row.typetext = sigText;
            row.parent = parentb;  row.depth = depthb;
            row.caller = whoKind;  row.callerref = whoRef;
            row.argcount = argcb;  row.args = argsText;
            // Exit row only: an entry row is written before the activation has an outcome.
            row.outcome = (std::strcmp(kind, "exit") == 0) ? OutcomeName(f.outcome) : "";
            row.ret = retText;  row.rettype = retType;
            row.ticks = ticksb;  row.tracerticks = tracerb;  row.trust = trustText;
            // The breakpoint count belongs with the duration it explains, so the exit row alone.
            char breaksb[16] = {};
            if (emit::csv::HasBreaksColumn() && std::strcmp(kind, "exit") == 0)
            {
                _snprintf_s(breaksb, _TRUNCATE, "%u", f.breaks);
                row.breaks = breaksb;
            }
            emit::csv::WriteRow(row);
        }

        // Counts a full table in the totals, which the table itself does not own.
        Proc* FindProc(std::uint64_t trailer)
        {
            Proc* p = g_procs.FindOrInsert(trailer);
            if (!p) Bump(g_totals.tableFull);
            return p;
        }

        // On every push, so a recompiled module is picked up at once.
        void ResolveInto(Proc* p, std::uint64_t trailer, ResolvedName* nm)
        {
            if (nm) { nm->qualModule[0] = 0; nm->function[0] = 0; }
            Identity id;
            if (Resolve(trailer, id) && id.function[0])
            {
                if (nm) _snprintf_s(nm->function, sizeof(nm->function),
                                    _TRUNCATE, "%s", id.function);
                // No project name: objTable+0x90 yields "ThisWorkbook", a document module.
                if (nm)
                {
                    if (id.workbook[0] && id.module[0])
                        _snprintf_s(nm->qualModule, sizeof(nm->qualModule),
                                    _TRUNCATE, "[%s]%s", id.workbook, id.module);
                    else if (id.module[0])
                        _snprintf_s(nm->qualModule, sizeof(nm->qualModule),
                                    _TRUNCATE, "%s", id.module);
                }
            }
            // A walk that did not check out leaves the row unnamed rather than
            // leaving whatever a previous call happened to write.

            // Compared and written under the flag readers take, so no row reads a
            // torn name. A loser skips; the next push publishes.
            if (p && nm && InterlockedCompareExchange(&p->publishing, 1, 0) == 0)
            {
                if (strcmp(p->function, nm->function) != 0 ||
                    strcmp(p->qualModule, nm->qualModule) != 0)
                {
                    _snprintf_s(p->function,   _TRUNCATE, "%s", nm->function);
                    _snprintf_s(p->qualModule, _TRUNCATE, "%s", nm->qualModule);
                }
                InterlockedExchange(&p->publishing, 0);
            }
        }

        std::uint64_t Now()
        {
            return static_cast<std::uint64_t>(core::QpcNow());
        }

        // xlfCaller answers from inside VBE7's interpreter: it is per-thread state and this is
        // the calculating thread. Its own guarded frame keeps a fault off the circuit breaker.

        bool AskCaller(core::Caller* out)
        {
            __try { core::ReadCaller(*out); return true; }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
        }

        bool ReadTrailer(std::uint64_t sp, std::uint64_t& out)
        {
            __try
            {
                out = *reinterpret_cast<std::uint64_t*>(sp + kTrailerOffset);
                return true;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
        }

        // Guarded though the block is ours: its offsets are an ABI shared with hand-written
        // assembly, which breaks silently.
        bool ReadReg(std::uint64_t savedRegs, int off, std::uint64_t& out)
        {
            if (!savedRegs) return false;
            __try
            {
                out = *reinterpret_cast<std::uint64_t*>(savedRegs + off);
                return true;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
        }

        ThreadState* State()
        {
            ThreadState* s = &t_state;
            const std::uint32_t gen = static_cast<std::uint32_t>(g_generation);
            // First use (generation 0, zero-filled by the loader) or a previous arming session.
            if (s->generation != gen)
            {
                s->depth = 0; s->current = 0; s->currentSp = 0; s->statements = 0;
                s->errActive = false;
                s->errSpan = 0;
                s->doEventsMarks = 0;
                s->generation = gen;
            }
            return s;
        }

        // Maxima are read-modify-write, so they need CAS where everything else
        // in Totals uses an interlocked add.
        void RaiseMax64(volatile std::uint64_t* target, std::uint64_t value)
        {
            volatile LONG64* t = reinterpret_cast<volatile LONG64*>(target);
            for (;;)
            {
                const LONG64 cur = *t;
                if (static_cast<std::uint64_t>(cur) >= value) return;
                if (InterlockedCompareExchange64(t, static_cast<LONG64>(value), cur) == cur)
                    return;
            }
        }
        void RaiseMax32(volatile std::uint32_t* target, std::uint32_t value)
        {
            volatile LONG* t = reinterpret_cast<volatile LONG*>(target);
            for (;;)
            {
                const LONG cur = *t;
                if (static_cast<std::uint32_t>(cur) >= value) return;
                if (InterlockedCompareExchange(t, static_cast<LONG>(value), cur) == cur)
                    return;
            }
        }

        // End and disarm aside, only an error unwind leaves without an epilogue, whether
        // Err.Raise or the runtime raised it. `ranEpilogue` is (exitOp != 0).
        void DecideOutcomeAtClose(ThreadState* s, Frame& f, bool ranEpilogue, bool stillRunning)
        {
            // Already resolved as the catcher -- unless an error then left it, which makes it
            // that error's thrower.
            if (f.outcome == kOutHandled && ranEpilogue) return;
            // `End` already settled this one, and no epilogue ran anywhere, so
            // every rule below would read it as an unwind it was not part of.
            if (f.outcome == kOutAbandoned) return;
            // Disarm closed it while it was still running: no epilogue, but no throw either.
            if (stillRunning) return;

            // Ran its own epilogue while a deeper throw was unclaimed, so it resumed
            // and caught it. Needed when the frame's next statement is its end.
            if (ranEpilogue && s->errActive && f.span < s->errSpan)
            {
                f.outcome    = kOutHandled;
                Bump(g_totals.handled);
                s->errActive = false;
                return;
            }

            // A same-frame On Error Resume Next of a real error also reads `returned`.
            if (ranEpilogue) return;

            // A deeper error already unwinding (span greater than this frame's) is the fatal
            // one, and this frame is unwound through. Otherwise the error started here.
            const bool deeperError = s->errActive && s->errSpan > f.span;

            if (!deeperError)
            {
                f.outcome     = kOutThrew;
                Bump(g_totals.threw);
                s->errActive  = true;
                s->errSpan    = f.span;
                return;
            }

            if (f.outcome == kOutReturned)
            {
                f.outcome = kOutUnwound;
                Bump(g_totals.unwound);
            }
        }

        // Class and form Functions all leave through 1664, so the store that wrote the result is
        // asked instead of the exit opcode. [R14-0x18] is otherwise an ordinary local, so only
        // these four Variant stores count there; two known stores of different types refuse it.
        bool IsVariantResultStore(std::uint16_t op)
        {
            return op == 694 || op == 707 || op == 702 || op == 703;   // FStVar, FStVarCopy, FStVarAd[Func]
        }

        RetKind DecideReturnKind(std::uint64_t trailer, std::uint16_t exitOp, std::uint16_t& storeOp)
        {
            storeOp = 0;
            RetKind kind = ExitReturnKind(exitOp);
            if (trailer == 0 || (kind != RetKind::Unknown && kind != RetKind::LongLongOrArray))
                return kind;

            // One scan serves both questions below.
            std::uint16_t ops[16]; std::int32_t offs[16];
            const int n = ReturnStoreCandidates(trailer, ops, offs, 16);

            if (kind == RetKind::Unknown)
            {
                // 671 only pushes the result slot's address: it is an array assignment when
                // nothing else writes the slot, and otherwise the function passing or updating
                // its own result (`F = F & x`), so a concrete store outranks it.
                RetKind fromStore = RetKind::Unknown;
                std::uint16_t fromOp = 0;
                bool conflict = false, addressOnly = false;
                for (int i = 0; i < n; ++i)
                {
                    if (offs[i] == -8 && ops[i] == 671) { addressOnly = true; continue; }
                    const RetKind k2 = (offs[i] == -0x18)
                                     ? (IsVariantResultStore(ops[i]) ? RetKind::Variant : RetKind::Unknown)
                                     : StoreReturnKind(ops[i]);
                    if (k2 == RetKind::Unknown) continue;   // not a store we know
                    if (fromStore == RetKind::Unknown) { fromStore = k2; fromOp = ops[i]; }
                    else if (fromStore != k2) conflict = true;
                }
                if (!conflict && fromStore != RetKind::Unknown)
                { kind = fromStore; storeOp = fromOp; }
                else if (!conflict && addressOnly)
                { kind = RetKind::LongLongOrArray; storeOp = 671; }
            }

            // Slot 634 serves LongLong and every typed array; the store to [R14-8] tells them
            // apart (699 FStI8, 671 array assign). With neither, the read is refused downstream.
            if (kind == RetKind::LongLongOrArray && storeOp == 0)
                for (int i = 0; i < n && storeOp == 0; ++i)
                    if (offs[i] == -8 && (ops[i] == 699 || ops[i] == 671)) storeOp = ops[i];
            return kind;
        }

        // The "after" for ByRef arguments: a ByRef slot points at the caller's variable, still
        // live at the exit opcode. ByVal is not reported, as the caller never sees that copy.
        // Returns `&out` when the arguments moved.
        const ArgCapture* ReadByRefChanges(const Frame& f, std::uint64_t r14, ArgCapture& out)
        {
            if (!(f.hasByRef && r14 != 0 && InterlockedCompareExchange(&g_capArgs, 0, 0) != 0))
                return nullptr;

            Bump(g_totals.byrefEligible);
            // Only the ByRef slots are re-read, rendered and compared.
            out.byRefOnly = true;
            CaptureArgs(f.trailer, r14, out);
            if (!out.ok)
            {
                Bump(g_totals.byrefDeclined);
                return nullptr;
            }
            // A shape says nothing about contents: comparing two would call any change "same".
            if (f.argsShapeOnly || out.byRefShapeOnly)
            {
                Bump(g_totals.byrefDeclined);
                return nullptr;
            }
            if (out.byRefHash != f.argsHash)
            {
                Bump(g_totals.byrefChanged);
                return &out;
            }
            Bump(g_totals.byrefSame);
            return nullptr;
        }

        // How the result read went, in the totals: read, declined, or never attempted because
        // the exit opcode has no mapping. The store candidates are recorded for that last case.
        void CountReturnOutcome(std::uint64_t trailer, std::uint64_t r14, RetKind kind,
                                std::uint16_t exitOp, bool haveRet)
        {
            const bool readable = (r14 != 0 && kind != RetKind::None && kind != RetKind::Unknown);
            if (haveRet)
                Bump(g_totals.returnsRead);
            else if (readable && InterlockedCompareExchange(&g_capRet, 0, 0) == 0)
                Bump(g_totals.returnsOff);
            else if (readable)
                Bump(g_totals.returnsDeclined);
            else if (r14 != 0 && kind == RetKind::Unknown && exitOp != 0)
            {
                g_unmappedExitOps.Note(exitOp);
                Bump(g_totals.returnsUnmapped);
                std::uint16_t ops[8]; std::int32_t offs[8];
                const int n = ReturnStoreCandidates(trailer, ops, offs, 8);
                for (int i = 0; i < n; ++i) g_retStores.Note(ops[i], offs[i]);
            }
        }

        // An error has left VBA. Cleared, or a later unrelated activation would close `unwound`.
        void EndEscapedError(ThreadState* s)
        {
            s->errActive = false;
            s->errSpan   = 0;
            Bump(g_totals.errEscaped);
        }

        // It ran off the bottom of the shadow stack: past every frame this thread
        // had, into a cell as #VALUE! or into whatever called the macro.
        void ClearErrorIfEscaped(ThreadState* s)
        {
            if (s->depth == 0 && s->errActive) EndEscapedError(s);
        }

        // It left through a frame Excel started to compute a cell. That frame reads
        // `unhandled`; the frame beneath, a macro that recalculated say, never sees it.
        void EndErrorAtExcelEntry(ThreadState* s, Frame& f)
        {
            if (!(f.fromExcel && s->errActive && s->errSpan >= f.span)) return;
            if (f.outcome != kOutAbandoned) f.outcome = kOutUnhandledErr;
            EndEscapedError(s);
        }

        // `r14` is the ending activation's frame base and comes only from the exit-opcode path:
        // on other close paths R14 belongs to another activation. `exitOp` names the return
        // kind, or 0.
        void CloseFrame(ThreadState* s, std::uint64_t nowTicks, std::uint64_t r14 = 0,
                        std::uint16_t exitOp = 0, std::int32_t exitOperand = 0,
                        const char* closedBy = "backstop", bool stillRunning = false)
        {
            if (s->depth <= 0) return;
            Frame& f = s->stack[s->depth - 1];

            // Match whole words: "end" and "exit" share a first letter.
            if (closedBy)
            {
                if      (std::strcmp(closedBy, "backstop") == 0)
                    Bump(g_totals.closedByBackstop);
                else if (std::strcmp(closedBy, "flush") == 0)
                    Bump(g_totals.closedByFlush);
                else if (std::strcmp(closedBy, "end") == 0)
                    Bump(g_totals.closedByEnd);
            }

            // exitOp != 0 means it ran its own epilogue, which an unhandled throw never does.
            DecideOutcomeAtClose(s, f, /*ranEpilogue=*/ exitOp != 0,
                                 stillRunning);
            EndErrorAtExcelEntry(s, f);

            Proc* p = FindProc(f.trailer);
            if (p)
            {
                InterlockedAdd64(reinterpret_cast<volatile LONG64*>(&p->ticks),
                                 static_cast<LONG64>(nowTicks - f.startTicks));
                InterlockedAdd64(reinterpret_cast<volatile LONG64*>(&p->statements),
                                 static_cast<LONG64>(s->statements - f.stmtsAtEntry));
            }
            // The result. The exit opcode says what kind it is and whether one exists; slot 634
            // alone needs the store opcode to split LongLong from an array.
            core::TextBuf& retText = t_retRender;
            retText.Clear();
            const char* retType = "";
            bool haveRet = false;
            std::uint16_t storeOp = 0;
            const RetKind kind = DecideReturnKind(f.trailer, exitOp, storeOp);
            const bool readable = (r14 != 0 && kind != RetKind::None && kind != RetKind::Unknown);
            if (readable && InterlockedCompareExchange(&g_capRet, 0, 0) != 0)
            {
                core::WithValueWriter(retText, [&](core::ValueWriter& w)
                { haveRet = DescribeReturnKind(r14, kind, storeOp, exitOperand, w, &retType); });
                haveRet = haveRet && !retText.Over();
            }
            // Reuses the per-thread render buffer: the entry's text is written and only its hash kept.
            ArgCapture outArgs;
            outArgs.text = &t_argRender;
            const ArgCapture* outArgsPtr = ReadByRefChanges(f, r14, outArgs);

            if (!g_topLevelOnly || f.shownDepth == 1)
                EmitRow("exit", f, nowTicks, nowTicks - f.startTicks, f.shownDepth, p,
                        outArgsPtr, nullptr, nullptr,
                        haveRet ? retText.Text() : "",
                        haveRet ? retType : "",
                        closedBy, /*argValuesOnly=*/true);
            CountReturnOutcome(f.trailer, r14, kind, exitOp, haveRet);
            --s->depth;
            Bump(g_totals.framesClosed);

            ClearErrorIfEscaped(s);
        }
    }

    // -------------------------------------------------------------------
    namespace
    {
        // Small on purpose: this runs on every VBA statement, so a repeating fault means stop
        // tracing rather than keep taking the risk on someone's live Excel.
        constexpr LONG kFaultsBeforeTrip = 8;

        // Tripped, every hook returns at once and the table stays patched but inert: unpatching
        // from a faulting thread would be a second way to crash. Nothing shared is reset until
        // inHook drains.
        struct Breaker
        {
            volatile LONG faults  = 0;
            volatile LONG tripped = 0;
            volatile LONG inHook  = 0;
        };
        Breaker g_breaker;

        void NoteHookFault()
        {
            const LONG n = InterlockedIncrement(&g_breaker.faults);
            Bump(g_totals.hookFaults);
            if (n >= kFaultsBeforeTrip && InterlockedExchange(&g_breaker.tripped, 1) == 0)
                g_totals.tripped = true;
        }

        // A stack overflow caught without restoring the guard page leaves the thread primed to
        // die later, so the handler calls _resetstkoflw and trips the breaker at once.
        __declspec(thread) unsigned t_lastExceptionCode = 0;


        // Every code is handled: vbathunk.asm's unwind info describes an empty prologue, so a
        // search would unwind to a garbage return address. Fix that first (.pushreg per push,
        // .setframe rbx) before declining any code.
        int HookFilter(unsigned code)
        {
            t_lastExceptionCode = code;
            return EXCEPTION_EXECUTE_HANDLER;
        }

        void HandleHookFault()
        {
            if (t_lastExceptionCode == static_cast<unsigned>(EXCEPTION_STACK_OVERFLOW))
            {
                Bump(g_totals.stackOverflows);
                _resetstkoflw();
                if (InterlockedExchange(&g_breaker.tripped, 1) == 0) g_totals.tripped = true;
            }
            // A fault while writing a row must not leave the file's lock held.
            emit::csv::ReleaseHeldByThisThread();
            NoteHookFault();
        }

        void OnStatementBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void OnBreakpointBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void OnStopBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void OnExitBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void CloseAtExit(std::uint64_t dispatchSp, std::uint64_t savedRegs, std::uint64_t now);
        void OnEndBody();
    }

    // The hook bodies are SEH-wrapped: a fault in the tracer's own code would otherwise
    // propagate into the VBA interpreter on a live calc thread. Wrapper plus body, because
    // __try may not share a frame with objects that need unwinding.
    namespace
    {
        using HookBody = void (*)(std::uint64_t, std::uint64_t);
        __declspec(thread) int t_hookDepth = 0;

        // The re-entry guard covers VBA run by a COM call made from inside a hook.
        void RunHook(HookBody body, std::uint64_t dispatchSp, std::uint64_t savedRegs)
        {
            if (g_breaker.tripped) return;
            if (t_hookDepth) { Bump(g_totals.hookReentries); return; }
            ++t_hookDepth;
            InterlockedIncrement(&g_breaker.inHook);
            __try { body(dispatchSp, savedRegs); }
            __except (HookFilter(GetExceptionCode())) { HandleHookFault(); }
            InterlockedDecrement(&g_breaker.inHook);
            --t_hookDepth;
        }

        void OnEndAdapter(std::uint64_t, std::uint64_t) { OnEndBody(); }
    }

    extern "C" void XRayVbaOnStatement(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnStatementBody, dispatchSp, savedRegs);
    }

    extern "C" void XRayVbaOnBreakpoint(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnBreakpointBody, dispatchSp, savedRegs);
    }

    extern "C" void XRayVbaOnStop(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnStopBody, dispatchSp, savedRegs);
    }

    extern "C" void XRayVbaOnExit(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnExitBody, dispatchSp, savedRegs);
    }

    // The End opcode has no frame to read; its arguments are unused.
    extern "C" void XRayVbaOnEnd(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnEndAdapter, dispatchSp, savedRegs);
    }

    namespace
    {
    // A statement running after a throw means its frame caught it. By then the unwind's closes
    // are done, so the top frame is the one that resumed; it must predate the throw, since a
    // destructor opened during the unwind has a higher span.
    void MarkHandlerIfErrorResumed(ThreadState* s)
    {
        if (!(s->errActive && s->depth > 0 &&
              s->stack[s->depth - 1].span < s->errSpan))
            return;

        Frame& h = s->stack[s->depth - 1];
        // A frame that threw and then caught its own error keeps `threw`: it is
        // the more informative of the two, and the chain is simply short.
        if (h.outcome != kOutThrew) h.outcome = kOutHandled;
        Bump(g_totals.handled);

        s->errActive = false;
    }

    // Asked before the fast path: a UDF that raises never reaches its epilogue, so its next
    // call arrives with the same trailer at the same rsp, the fast path's exact condition.
    bool IsActivationStart(std::uint64_t trailer, std::uint64_t savedRegs)
    {
        switch (ActivationStart(trailer, savedRegs))
        {
        case Activation::Yes:
            Bump(g_totals.ipEntries);
            return true;
        case Activation::Unavailable:
            Bump(g_totals.ipUnavailable);
            return false;
        case Activation::No:
            break;
        }
        return false;
    }

    // Counting only, so the fast path can be seen working.
    void NoteStackPointerMove(ThreadState* s, std::uint64_t trailer, std::uint64_t dispatchSp)
    {
        if (!(trailer == s->current && s->currentSp != 0)) return;
        if (dispatchSp < s->currentSp)
            Bump(g_totals.sameTrailerDeeper);
        else if (dispatchSp > s->currentSp)
            Bump(g_totals.sameTrailerShallower);
    }

    // The stack grows down, so a frame whose stack was released has a smaller sp.
    void CloseFramesThatReturned(ThreadState* s, std::uint64_t dispatchSp, std::uint64_t now)
    {
        while (s->depth > 0 && s->stack[s->depth - 1].sp < dispatchSp)
            CloseFrame(s, now);
    }

    // A VBA call keeps its caller's cell, so a cell that differs from the frame beneath's is a
    // fresh worksheet-function entry, whose unhandled error Excel turns into #VALUE!. A sheet
    // event has no calling cell, so its unhandled error reads `threw`.
    bool IsExcelTheCaller(ThreadState* s, bool callerIsCell, std::uint64_t callerHash)
    {
        if (!callerIsCell) return false;
        return s->depth < 2 || s->stack[s->depth - 2].callerHash != callerHash;
    }

    // Excel runs timer macros and events from inside DoEvents, so such a procedure is not the
    // waiting frame's callee. A stale mark (End inside DoEvents runs no epilogue) is dropped.
    bool WaitingInDoEvents(ThreadState* s)
    {
        while (s->doEventsMarks > 0 && s->doEventsAt[s->doEventsMarks - 1] > s->depth)
            --s->doEventsMarks;
        return s->doEventsMarks > 0 && s->doEventsAt[s->doEventsMarks - 1] == s->depth;
    }

    // Is this frame still executing? The interpreter keeps a frame's trailer at the dispatch
    // rsp recorded for it, so a returned frame's slot holds something else. Needed because an
    // unhandled error runs no epilogue and the backstop may never fire.
    bool FrameStillLive(const Frame& f)
    {
        std::uint64_t onStack = 0;
        return ReadTrailer(f.sp, onStack) && onStack == f.trailer;
    }

    void CloseFramesThatLostTheirStack(ThreadState* s, std::uint64_t now)
    {
        while (s->depth > 0)
        {
            if (FrameStillLive(s->stack[s->depth - 1])) return;
            CloseFrame(s, now);
        }
    }

    // Only a prologue opens a frame, since `GoSub` moves rsp within one activation. A procedure
    // already running at arm has no prologue, so it is opened anyway as `lateOpen`.
    bool ShouldOpenFrame(ThreadState* s, std::uint64_t trailer, bool atEntry,
                         std::uint64_t dispatchSp, bool& lateOpen)
    {
        bool onOurStack = false;
        for (int i = 0; i < s->depth; ++i)
            if (s->stack[i].trailer == trailer) { onOurStack = true; break; }

        lateOpen = !atEntry && !onOurStack;
        if (!atEntry && !lateOpen)
        {
            // Same activation, different interpreter rsp: GoSub, Return, or an
            // error handler moving the stack inside one procedure.
            if (s->depth == 0 || s->stack[s->depth - 1].sp != dispatchSp)
                Bump(g_totals.ipIntraProc);
        }
        return atEntry || lateOpen;
    }

    // Once per activation, as xlfCaller is a call into Excel. Always asked: the cell is needed
    // to place an error that escapes into it.
    void IdentifyCaller(core::Caller& who)
    {
        if (!AskCaller(&who))
        {
            Bump(g_totals.callerFaults);
            who.kind[0] = 0; who.desc[0] = 0; who.isCell = false;
        }
        else if (who.isCell)
            Bump(g_totals.callerCell);
        else if (std::strcmp(who.kind, "unavailable") == 0)
            Bump(g_totals.callerUnavailable);
        else
            Bump(g_totals.callerOther);
    }

    // Order matters: stale-close before push, and arguments captured at the first statement,
    // before the body can overwrite a ByRef.
    void OpenFrame(ThreadState* s, std::uint64_t trailer, std::uint64_t savedRegs,
                   std::uint64_t dispatchSp, bool atEntry, bool lateOpen,
                   std::uint64_t now)
    {
        if (lateOpen)
            Bump(g_totals.ipLateOpen);

        // A frame still open at this rsp cannot be running: this statement took its place. The
        // same procedure there is the frame a raising UDF left behind.
        if (s->depth > 0 && s->stack[s->depth - 1].sp == dispatchSp)
        {
            if (atEntry && s->stack[s->depth - 1].trailer == trailer)
                Bump(g_totals.ipStaleClosed);
            CloseFrame(s, now);
        }

        // A frame that cannot be pushed would leave every later exit mismatched, so the tracer
        // stands down rather than trace a partial tree.
        if (!s->stack.Reserve(s->depth))
        {
            Bump(g_totals.stackGrowFailures);
            if (InterlockedExchange(&g_breaker.tripped, 1) == 0) g_totals.tripped = true;
            return;
        }

        for (int i = 0; i < s->depth; ++i)
            if (s->stack[i].trailer == trailer)
            {
                Bump(g_totals.recursions);
                break;
            }

        // Excel ran this from inside the open frame's DoEvents: a new chain, not a callee.
        const bool newChain = atEntry && s->depth > 0 && WaitingInDoEvents(s);
        if (newChain) Bump(g_totals.doEventsChains);
        const std::uint64_t parent = (s->depth > 0 && !newChain) ? s->stack[s->depth - 1].span : 0;
        const int shownDepth = (s->depth > 0 && !newChain) ? s->stack[s->depth - 1].shownDepth + 1 : 1;
        Frame& f = s->stack[s->depth++];
        f.trailer      = trailer;
        f.startTicks   = now;
        f.tracerAtEntry = core::TracerTicks();
        // -1 because the statement that opened this frame is already counted; otherwise the
        // callee's first statement is billed to its caller.
        f.stmtsAtEntry = s->statements - 1;
        f.sp           = dispatchSp;
        f.span         = emit::csv::NextSpan();
        f.parent       = parent;
        f.shownDepth   = shownDepth;
        f.outcome      = kOutReturned;
        f.breaks       = 0;
        f.hasByRef     = false;
        f.argsHash     = 0;
        f.argsShapeOnly = false;
        f.fromExcel    = false;
        f.callerHash   = 0;

        Bump(g_totals.framesOpened);

        ResolvedName nm{};
        Proc* p = FindProc(trailer);
        if (p)
        {
            ResolveInto(p, trailer, &nm);   // fresh every call -- never cached
            Bump(p->calls);
            RaiseMax64(&p->maxDepth, static_cast<std::uint64_t>(s->depth));
        }
        // R14 is the VBA frame base. Checked, so a broken assembly contract is not mistaken for
        // a frame with no base.
        std::uint64_t r14 = 0;
        const bool haveR14 = ReadReg(savedRegs, kReg_r14, r14);
        if (!haveR14)
            Bump(g_totals.regReadFailures);

        ArgCapture args;
        args.text = &t_argRender;
        // Off, the decline is counted, so an empty column never reads as "no arguments".
        if (haveR14 && InterlockedCompareExchange(&g_capArgs, 0, 0) != 0)
        {
            CaptureArgs(trailer, r14, args);
            // Kept so the exit can tell whether anything moved.
            if (args.ok)
            {
                // Either test is enough: the signature names a ByRef type the p-code declared,
                // while `viaPointer` catches a slot dereferenced to read its value though
                // nothing declared it, such as a write-only String.
                f.hasByRef = SignatureHasByRef(args.signature) || args.viaPointer;
                f.argsHash = args.byRefHash;
                f.argsShapeOnly = args.byRefShapeOnly;
            }
        }

        core::Caller who;      // NOT {} -- IdentifyCaller fills it on every path
        IdentifyCaller(who);
        f.callerHash = who.isCell ? core::Fnv1aText(who.desc) : 0;
        f.fromExcel  = IsExcelTheCaller(s, who.isCell, f.callerHash);

        if (!g_topLevelOnly || f.shownDepth == 1)
            EmitRow("entry", f, now, 0, f.shownDepth, p, &args, &nm, &who);

        RaiseMax32(&g_totals.maxDepth, static_cast<std::uint32_t>(s->depth));
    }

    void OnStatementBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        ThreadState* s = State();
        if (!s) return;

        ++s->statements;
        Bump(g_totals.statements);

        // Order matters: was an error caught, where are we, is this the same frame, what has
        // returned, does a frame open.
        MarkHandlerIfErrorResumed(s);

        std::uint64_t trailer = 0;
        // Zero is no procedure: nothing to name, count or open a frame for.
        if (!ReadTrailer(dispatchSp, trailer) || trailer == 0)
        {
            Bump(g_totals.faults);
            return;
        }

        const bool atEntry = IsActivationStart(trailer, savedRegs);

        // Fast path: same procedure at the same rsp is the next statement of the current frame.
        if (!atEntry && trailer == s->current && dispatchSp == s->currentSp)
        {
            Bump(g_totals.sameTrailerSameSp);
            return;
        }

        NoteStackPointerMove(s, trailer, dispatchSp);

        const std::uint64_t now = Now();
        Bump(g_totals.transitions);

        // (trailer, sp) identifies an activation: the trailer names only a procedure, and the
        // dispatch rsp falls one step per VBA call and is stable within a frame.

        CloseFramesThatReturned(s, dispatchSp, now);
        if (atEntry) CloseFramesThatLostTheirStack(s, now);

        bool lateOpen = false;
        if (ShouldOpenFrame(s, trailer, atEntry, dispatchSp, lateOpen))
            OpenFrame(s, trailer, savedRegs, dispatchSp, atEntry, lateOpen, now);

        s->current   = trailer;
        s->currentSp = dispatchSp;
        core::TracerTicks() += Now() - now;
    }

    // A breakpointed statement is a statement first: it opens or continues its frame exactly as
    // BoS would, since the editor is about to stop on it and BoS's handler never runs. The stop
    // is then charged to the frame the statement belongs to.
    void OnBreakpointBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        OnStatementBody(dispatchSp, savedRegs);
        Bump(g_totals.breakpointStops);
        ThreadState* s = State();
        if (s && s->depth > 0 && s->stack[s->depth - 1].sp == dispatchSp)
            ++s->stack[s->depth - 1].breaks;
    }

    // A `Stop` statement, charged to the frame it is in. It touches no frame state: its
    // statement's BoS has already run, and the editor resumes with the frame intact.
    void OnStopBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        (void)savedRegs;
        Bump(g_totals.stopStatements);
        ThreadState* s = State();
        // The same frame match the breakpoint hook uses: both dispatch at the interpreter rsp the
        // frame was opened at.
        if (s && s->depth > 0 && s->stack[s->depth - 1].sp == dispatchSp)
            ++s->stack[s->depth - 1].breaks;
    }

    // The exit opcode fires exactly once per activation; rsp stays the backstop, since an error
    // unwind and `End` fire none. Matching both trailer and sp prevents a double close.
    void OnExitBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        // Stamped first, so the hook's own reads fall outside the activation it ends.
        const std::uint64_t now = Now();
        CloseAtExit(dispatchSp, savedRegs, now);
        core::TracerTicks() += Now() - now;
    }

    void CloseAtExit(std::uint64_t dispatchSp, std::uint64_t savedRegs, std::uint64_t now)
    {
        Bump(g_totals.exits);

        // A failed read here is counted, not swallowed.
        std::uint64_t trailer = 0, r14 = 0;
        if (!ReadTrailer(dispatchSp, trailer) || trailer == 0)
        {
            Bump(g_totals.faults);
            return;
        }
        if (!ReadReg(savedRegs, kReg_r14, r14))
            Bump(g_totals.regReadFailures);

        // Does this exit end the procedure? A GoSub's `Return` fires one mid-activation
        // (vbaboundary.h). Both non-Yes answers close nothing and are counted apart; a missed
        // close is corrected later, a wrong close writes a row for a call VBA never made.
        std::uint16_t exitOp = 0;
        const Ending ending = ProcedureEnd(savedRegs, &exitOp);
        if (exitOp != 0) g_exitOpsSeen.Note(exitOp);
        if (ending != Ending::Yes)
        {
            Bump(ending == Ending::Unreadable ? g_totals.exitOpUnreadable : g_totals.exitNotProcedureEnd);
            return;
        }

        ThreadState* s = State();
        if (!s) return;
        if (s->depth <= 0)
        {
            Bump(g_totals.unmatchedExits);
            return;
        }

        const Frame& top = s->stack[s->depth - 1];
        if (top.trailer == trailer && top.sp == dispatchSp)
        {
            std::int32_t exitOperand = 0;
            if (ExitReturnKind(exitOp) == RetKind::RecordInFrame && !ExitOperand(savedRegs, exitOperand))
                exitOperand = 0;   // refused downstream
            CloseFrame(s, now, r14, exitOp, exitOperand, "exit");
            // Forget the fast-path cache, or the next activation of the same procedure at
            // the same rsp matches it and never opens a frame.
            s->current   = 0;
            s->currentSp = 0;
            Bump(g_totals.exitClosed);
        }
        else
        {
            Bump(g_totals.exitNoMatch);
        }
    }
        // `End` fires no exit opcode for any frame it kills, and the backstop cannot close them
        // when the next chain starts deeper: rsp cannot tell "nested" from "abandoned".
        // This thread only, because the shadow stack is TLS.
        void OnEndBody()
        {
            ThreadState* s = State();
            if (!s) return;
            const std::uint64_t now = Now();
            while (s->depth > 0)
            {
                s->stack[s->depth - 1].outcome = kOutAbandoned;
                CloseFrame(s, now, 0, 0, 0, "end");
            }
            // Nothing of the dead chain may be matched against the next one.
            s->current = 0; s->currentSp = 0;
            s->errActive = false; s->errSpan = 0;
        }

    }   // anonymous namespace

    // -------------------------------------------------------------------

    // A macro run inside DoEvents may call DoEvents itself; past the small stack the innermost
    // marks are dropped, costing only the parenting of a chain nested that deep.
    void NoteDoEventsEnter()
    {
        ThreadState* s = State();
        if (s->doEventsMarks < static_cast<int>(sizeof(s->doEventsAt) / sizeof(s->doEventsAt[0])))
            s->doEventsAt[s->doEventsMarks++] = s->depth;
    }

    void NoteDoEventsLeave()
    {
        ThreadState* s = State();
        if (s->doEventsMarks > 0) --s->doEventsMarks;
    }

    void ReleaseThreadState() { t_state.stack.Release(); }

    void FlushOpenFrames()
    {
        // t_state is the storage itself, not a pointer to it, so a thread that
        // never traced has generation 0 and depth 0 and this is a no-op for it.
        ThreadState* s = &t_state;
        const std::uint64_t now = Now();
        const std::uint64_t here = reinterpret_cast<std::uint64_t>(_AddressOfReturnAddress());
        bool running = false;
        while (s->depth > 0)
        {
            const Frame& top = s->stack[s->depth - 1];
            // Running if above our own stack with its trailer in place (a leftover from an
            // earlier unwind fails that); a frame beneath a running one runs too.
            running = running || (top.sp > here && FrameStillLive(top));
            CloseFrame(s, now, 0, 0, 0, "flush", running);
        }
        s->current = 0; s->currentSp = 0;
    }

    bool WaitForHooksQuiet(unsigned milliseconds)
    {
        const ULONGLONG deadline = GetTickCount64() + milliseconds;
        while (InterlockedCompareExchange(&g_breaker.inHook, 0, 0) != 0)
        {
            if (GetTickCount64() >= deadline) return false;
            Sleep(1);
        }
        return true;
    }

    int ProcTableSize() { return g_procs.Size(); }

    bool ProcSnapshot(int slot, unsigned long long& trailer, unsigned long long& calls,
                      char* module, int modCap, char* function, int fnCap)
    {
        Proc* pp = g_procs.At(slot);
        if (!pp) return false;
        Proc& p = *pp;
        const std::uint64_t tr = p.trailer;
        if (tr == 0) return false;                    // slot never used
        trailer = tr;
        calls   = p.calls;

        module[0] = 0; function[0] = 0;
        // A name being republished could tear; the trailer is always true.
        if (InterlockedCompareExchange(&p.publishing, 1, 0) == 0)
        {
            strncpy_s(module,   modCap, p.qualModule, _TRUNCATE);
            strncpy_s(function, fnCap,  p.function,   _TRUNCATE);
            InterlockedExchange(&p.publishing, 0);
        }
        return true;
    }

    void SetCapture(bool args, bool retval)
    {
        InterlockedExchange(&g_capArgs, args   ? 1 : 0);
        InterlockedExchange(&g_capRet,  retval ? 1 : 0);
    }

    void SetEmitTopLevelOnly(bool topOnly)
    {
        InterlockedExchange(&g_topLevelOnly, topOnly ? 1 : 0);
    }

    void ResetTracing()
    {
        g_procs.Reset();
        g_totals = Totals{};
        ResetIdentityCounts();
        ResetArgCounts();
        // Or p-code declines, including `Desynced`, accumulate for the life of the process.
        ResetPcodeCounts();
        // Arming is a deliberate human act, so the tracer gets another go from zero faults.
        g_breaker.faults = 0;
        InterlockedExchange(&g_breaker.tripped, 0);
        // Or they report opcodes from a previous arm as if this one had seen them.
        g_unmappedExitOps.Reset();
        g_exitOpsSeen.Reset();
        g_retStores.Reset();
        InterlockedIncrement(&g_generation);   // every thread drops its stale stack
    }

    SeenTable& UnmappedExitOps() { return g_unmappedExitOps; }
    SeenTable& RetStores()       { return g_retStores; }
    SeenTable& ExitOpsSeen()     { return g_exitOpsSeen; }
    ProcTable& Procs()           { return g_procs; }

    Totals ReadTotals()
    {
        Totals t = g_totals;
        t.procedures = static_cast<std::uint64_t>(g_procs.Count());
        return t;
    }
}
