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
        // kTrailerOffset (trailer pointer at [dispatchSp + 0xB8]) is in vbatrailer.h.


        // The saved-register block's layout is in vbaregs.h, shared with everything
        // that reads it.

        // ---- EXIT OPCODES WE CANNOT NAME A RETURN KIND FOR --------------
        //
        // Recorded by VALUE, because the value is the thing that has to be mapped; a
        // count alone would say a gap exists without saying which one to close.
        SeenTable g_unmappedExitOps;

        // The (store opcode, operand) pairs seen on procedures whose return
        // could not be typed, so the mapping is extended from evidence rather
        // than from a guess about what a String store looks like.
        SeenTable g_retStores;          // (store opcode, operand offset)

        // EVERY exit opcode seen, mapped or not. ProcedureEnd uses a DENY-LIST,
        // so an unknown mid-procedure exit construct would still close wrongly
        // (vbaboundary.h) -- this is how one gets named.
        SeenTable g_exitOpsSeen;

        const std::uint8_t kOutReturned = 0;
        const std::uint8_t kOutThrew    = 1;
        const std::uint8_t kOutUnwound  = 2;
        const std::uint8_t kOutHandled  = 3;
        // `End` kills the whole VBA session without an epilogue anywhere, so
        // these frames neither returned nor threw. Writing `returned` for them
        // would be a confident wrong cell -- the one thing this project will
        // not do -- so the abrupt end gets a name of its own.
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
            std::uint64_t stmtsAtEntry;
            std::uint64_t sp;            // interpreter rsp when this frame was seen
            std::uint64_t span;          // pairs this frame's entry and exit rows
            std::uint64_t parent;        // the span we were called from, 0 at the top
            // The depth rows report: 1 for a chain Excel started inside another's DoEvents,
            // though that chain still sits above it on the shadow stack.
            int           shownDepth;
            // HOW THIS ACTIVATION ENDED.
            //
            //   returned  it finished normally
            //   threw     an error left it, and it was the innermost frame to go
            //   unwound   the error passed THROUGH it -- it ran nothing after the
            //             error, so it did not handle it
            //   handled   it ran again after the error, so it caught it
            //
            // A chain reads threw -> unwound -> unwound -> handled from the throwing
            // frame outwards: where was it thrown, and who caught it.
            std::uint8_t  outcome;

            // EXCEL STARTED THIS FRAME as a worksheet-function activation, so an error
            // leaving it becomes a cell's #VALUE! rather than a VBA caller's error.
            bool          fromExcel;
            // The cell xlfCaller named for this frame, hashed; 0 when not a cell.
            // A frame whose cell differs from the one beneath is a fresh Excel entry.
            std::uint64_t callerHash;

            // True when the recovered signature has a ByRef parameter. Only
            // then is the exit re-read worth its cost, and only then can an
            // "after" value mean anything.
            bool          hasByRef;

            // A HASH of the ByRef slots, not the text: the text is up to 64 KB and the shadow
            // stack is fixed-size TLS, so keeping it per frame would cost
            // 64 KB x kMaxDepth to answer a yes/no question. The "before" text is
            // already in the entry ROW, which is where a reader gets it.
            std::uint64_t argsHash;
            // A ByRef argument was written as its shape alone at entry, so no exit can compare.
            bool          argsShapeOnly;

            // No caller gate: the exit opcode says whether a result exists (a
            // Sub leaves through slot 635) and of what kind, so the return is
            // read at every depth and whatever the caller was.
        };

        // A plain aggregate with no member initialisers, so it can live in thread-local storage
        // with no dynamic initialiser. `generation` starting at zero is load-bearing:
        // g_generation starts at 1, so the check below initialises the block on first use.
        struct ThreadState
        {
            Frame  stack[kMaxDepth];
            int    depth;
            std::uint64_t current;      // trailer of the running procedure
            std::uint64_t currentSp;    // and the rsp we last saw it at
            std::uint64_t statements;   // this thread's statement count
            // Activations the shadow stack had no room for, kept so the DEPTH
            // is still reportable: depth + overflowDepth is how deep VBA is.
            int    overflowDepth;
            std::uint32_t generation;   // which arming session this belongs to

            // An error in flight on this thread, set when the thrower closes. A frame that
            // ran nothing since unwound; a statement running again in a frame that predates
            // the throw means that frame caught it.
            bool          errActive;
            // DoEvents open on this thread, and the frame depth each call was made at, so a
            // procedure Excel runs while a frame waits there can be told from one it called.
            int           doEventsMarks;
            int           doEventsAt[8];

            // THE SPAN OF THE FRAME THAT THREW. Spans only ever increase, so a frame
            // with a LOWER span existed before the throw and can be the handler; a higher
            // one opened during the unwind and cannot.
            std::uint64_t errSpan;
        };

        // TLS, because hooks run reentrantly across threads. The STORAGE is
        // thread-local, not a thread-local pointer to the heap.
        __declspec(thread) ThreadState t_state;

        // The argument and return columns' buffers: per thread, on the heap, so ArgCapture
        // stays tiny on the deeply nested interpreter stack. Leaked at thread exit.
        __declspec(thread) core::TextBuf t_argRender;
        __declspec(thread) core::TextBuf t_retRender;

        // The procedure name/count table (Proc + the open-addressed map) lives in
        // vbaproctable.h/.cpp now. Its per-Proc counters are still incremented
        // here through the returned Proc*; ResolveInto (below) fills the names.
        ProcTable g_procs;
        Totals    g_totals;

        // Bumped on every ResetTracing: a thread cannot clear another thread's
        // TLS, so each notices its state is from a previous arming session and
        // clears its own. Without it the shadow stack carries frames across a
        // disarm/re-arm.
        volatile LONG g_generation = 1;

        // One "the shadow stack ran out" row per arming session. Reset by
        // ResetTracing along with everything else.
        volatile LONG g_cappedNoted = 0;

        // Emit rows only for depth-1 frames. Latched at arm; the totals count
        // every frame regardless, so the trace thins but the accounting does
        // not. The depth-capped marker row is NOT filtered: a truncation must
        // never be hidden by a mode.
        volatile LONG g_topLevelOnly = 0;
        // Argument capture can be turned off as a control. It is the only part of the hook that
        // follows pointers out of the frame and walks bytecode, so it has the largest fault
        // surface. Latched at arm; a disabled decode is counted, never silently blank.
        volatile LONG g_capArgs = 1;

        // Does the signature contain a ByRef parameter? ByRef is VBA's default, so `Sub
        // Calc(result As Double)` filling in `result` is ordinary code. The marker is the "&"
        // the type table uses.
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

        // The same trace file as the XLL side, and one sequence across both. A row is stamped
        // (`qpc`) when its event happened and numbered (`seq`) when it is written. Runs on a
        // frame change, never per statement.
        void EmitRow(const char* kind, const Frame& f, std::uint64_t qpc,
                     std::uint64_t durationTicks, int depth, Proc* p,
                     const ArgCapture* args = nullptr,
                     const ResolvedName* nm = nullptr,
                     const core::Caller* who = nullptr,
                     const char* retText = "", const char* retType = "",
                     const char* closedBy = "",
                     // EXIT ROWS CARRY THE VALUES ONLY. `argcount` and
                     // `typetext` are properties of the SIGNATURE, already on
                     // the entry row this pairs with by span; repeating them
                     // would break the row model's "entry rows only".
                     bool argValuesOnly = false)
        {
            // The only part that touches a FILE. Off, the hooks still run and
            // the totals still count, so a crash that survives this is not the
            // emitter.
            if (!emit::csv::IsOpen()) return;
            char spanb[24], tidb[16], qpcb[24], trailerb[32], ticksb[24];
            char parentb[24], depthb[16];
            _snprintf_s(spanb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(f.span));
            _snprintf_s(tidb, _TRUNCATE, "%lu", GetCurrentThreadId());
            _snprintf_s(qpcb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(qpc));
            _snprintf_s(trailerb, _TRUNCATE, "0x%llX",
                        static_cast<unsigned long long>(f.trailer));
            // `parent` is a span, so a missing row shows as a reference to a span that is not
            // in the file, where rebuilding the tree from depth and row order would attach
            // everything after a gap to the wrong caller.
            _snprintf_s(parentb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(f.parent));
            _snprintf_s(depthb, _TRUNCATE, "%d", depth);
            // On the exit row only. `returned` is written explicitly, so an error unwind never
            // looks like a clean return.
            //
            // `trust` says what closed the frame, which decides whether `ticks` is a
            // measurement or an upper bound: only the exit opcode fires at the moment an
            // activation ends; the backstop fires at the next statement and the flush at
            // disarm. A fully unhandled unwind fires no exit opcodes at all.
            const char* trustText = "";
            if (durationTicks)
            {
                _snprintf_s(ticksb, _TRUNCATE, "%llu",
                            static_cast<unsigned long long>(durationTicks));
                trustText = (closedBy && closedBy[0]) ? closedBy : "backstop";
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

            // `proc` is always the trailer, which tells two same-named procedures apart.
            // `module` is empty when unresolvable; `function` falls back to the trailer in hex.
            // The caller's stack-local resolution is preferred over the shared Proc entry,
            // which is rewritten on every push.
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
            // OUTCOME IS ITS OWN COLUMN and belongs to the exit row alone: an
            // entry row is written before the activation has an outcome, and
            // saying anything there would be a guess about the future.
            row.outcome = (std::strcmp(kind, "exit") == 0) ? OutcomeName(f.outcome) : "";
            row.ret = retText;  row.rettype = retType;
            row.ticks = ticksb;  row.trust = trustText;
            emit::csv::WriteRow(row);
        }

        // Find the Proc for a trailer, counting a full table in the totals (the
        // one totals touch the table itself does not own).
        Proc* FindProc(std::uint64_t trailer)
        {
            Proc* p = g_procs.FindOrInsert(trailer);
            if (!p) Bump(g_totals.tableFull);
            return p;
        }

        // On every frame push, so an edited-and-recompiled module is picked up
        // immediately; see Proc for why nothing here is cached.
        void ResolveInto(Proc* p, std::uint64_t trailer, ResolvedName* nm)
        {
            if (nm) { nm->qualModule[0] = 0; nm->function[0] = 0; }
            Identity id;
            if (Resolve(trailer, id) && id.function[0])
            {
                if (nm) _snprintf_s(nm->function, sizeof(nm->function),
                                    _TRUNCATE, "%s", id.function);
                // The project name is left out: objTable+0x90 yields "ThisWorkbook" on this
                // build, a document module, not the project.
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

        // Asking Excel from inside VBE7's interpreter, not an XLL callback. xlfCaller answers
        // because it is per-thread state and this is the calculating thread. Guarded in its own
        // leaf frame, so a fault here is not attributed to the hook and does not trip the
        // circuit breaker.

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

        // Guarded even though the block is ours and cannot fault today: the
        // offsets into it are an ABI shared with hand-written assembly, which is
        // exactly the kind of agreement that breaks silently.
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
            // Covers BOTH "first use on this thread" (generation 0, zero-filled
            // by the loader) and "left over from a previous arming session".
            if (s->generation != gen)
            {
                s->depth = 0; s->current = 0; s->currentSp = 0; s->statements = 0;
                s->overflowDepth = 0;
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

        // Decide the outcome as the frame closes, from how it left. Only an error unwind leaves
        // without an epilogue once End and disarm are set aside, whether Err.Raise or the
        // runtime (a division by zero) raised it. `ranEpilogue` is (exitOp != 0).
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

            // A DEEPER error already unwinding (span greater than this frame's) is the fatal
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

        // What kind of result this procedure left behind. The exit opcode says, for most
        // procedures (vbaretdecode.h); class and form Functions all leave through 1664, so the
        // instruction that wrote the result is asked instead.
        //
        // The offset identifies a Variant, not the opcode: a Variant result lives at [R14-0x18]
        // and is stored through different opcodes depending on what it holds. Two known stores
        // of different types cannot both be the result, so the type is then refused.
        RetKind DecideReturnKind(std::uint64_t trailer, std::uint16_t exitOp, std::uint16_t& storeOp)
        {
            storeOp = 0;
            RetKind kind = ExitReturnKind(exitOp);
            if (trailer == 0 || (kind != RetKind::Unknown && kind != RetKind::LongLongOrArray))
                return kind;

            // ONE SCAN serves both questions below.
            std::uint16_t ops[16]; std::int32_t offs[16];
            const int n = ReturnStoreCandidates(trailer, ops, offs, 16);

            if (kind == RetKind::Unknown)
            {
                RetKind fromStore = RetKind::Unknown;
                std::uint16_t fromOp = 0;
                bool conflict = false;
                for (int i = 0; i < n; ++i)
                {
                    const RetKind k2 = (offs[i] == -0x18) ? RetKind::Variant
                                                          : StoreReturnKind(ops[i]);
                    if (k2 == RetKind::Unknown) continue;   // not a store we know
                    if (fromStore == RetKind::Unknown) { fromStore = k2; fromOp = ops[i]; }
                    else if (fromStore != k2) conflict = true;
                }
                if (!conflict && fromStore != RetKind::Unknown)
                { kind = fromStore; storeOp = fromOp; }
            }

            // Slot 634 serves both a LongLong and every typed array; the store
            // that wrote [R14-8] tells them apart (699 = FStI8, 671 = array
            // assign). With neither in evidence the read is refused downstream
            // rather than resolved by whether the bytes look like a pointer.
            if (kind == RetKind::LongLongOrArray && storeOp == 0)
                for (int i = 0; i < n && storeOp == 0; ++i)
                    if (offs[i] == -8 && (ops[i] == 699 || ops[i] == 671)) storeOp = ops[i];
            return kind;
        }

        // ByRef arguments the procedure changed. Entry captured the arguments before the body
        // could touch them; this is the "after". A ByRef slot points at the caller's variable,
        // whose frame is still live at the exit opcode. ByVal is not reported: the caller never
        // sees that copy. Returns `&out` when the arguments moved.
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

        // ---- AN ERROR HAS LEFT VBA ---------------------------------------------
        //
        // The chain is over. Clear it so it cannot be read into a later, unrelated
        // activation, which would otherwise close stamped `unwound`.
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

        // Closes the frame on top. `r14` is the frame base of the activation that is ending,
        // and comes only from the exit-opcode path: on the other close paths R14 belongs to a
        // different activation. `exitOp` is the typed exit that names the return kind, or 0.
        void CloseFrame(ThreadState* s, std::uint64_t nowTicks, std::uint64_t r14 = 0,
                        std::uint16_t exitOp = 0, const char* closedBy = "backstop",
                        bool stillRunning = false)
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

            // The outcome is decided from HOW the frame left: exitOp != 0 means
            // it ran its own epilogue, which a real unhandled throw never does.
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
                { haveRet = DescribeReturnKind(r14, kind, storeOp, w, &retType); });
                haveRet = haveRet && !retText.Over();
            }
            // The "after" for the entry's "before": see ReadByRefChanges. Reuses
            // the same per-thread render buffer -- the entry's text is already
            // written and only its hash is kept, so nothing live is overwritten.
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
        // The threshold is small on purpose: this path runs on every VBA
        // statement in the process, so a fault that repeats is not a blip, and
        // the correct response to "the tracer is broken" is to stop tracing
        // rather than keep taking the risk on somebody's live Excel.
        constexpr LONG kFaultsBeforeTrip = 8;

        // THE CIRCUIT BREAKER, as one thing: faults seen, whether it has
        // TRIPPED (every hook then returns at once, leaving the dispatch table
        // patched but INERT -- unpatching from a faulting thread would be a
        // second way to crash), and how many threads are inside a hook right
        // now, since nothing shared may be reset until that drains.
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

        // A stack overflow is not an ordinary fault: catching one without restoring the guard
        // page leaves the thread primed to die later. The filter remembers the code, and the
        // handler calls _resetstkoflw and opens the breaker at once.
        __declspec(thread) unsigned t_lastExceptionCode = 0;


        // Every code is handled here, and that is load-bearing: vbathunk.asm's unwind info
        // describes an empty prologue, so EXCEPTION_CONTINUE_SEARCH would unwind through it to
        // a garbage return address. To decline a code, fix the thunk's unwind info first
        // (.pushreg per push and .setframe rbx).
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
        void OnExitBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void OnEndBody();
    }

    // The hook bodies are SEH-wrapped: a fault in the tracer's own code would otherwise
    // propagate into the VBA interpreter on a live calc thread. Wrapper plus body, because
    // __try may not share a frame with objects that need unwinding.
    namespace
    {
        using HookBody = void (*)(std::uint64_t, std::uint64_t);
        __declspec(thread) int t_hookDepth = 0;

        // Every hook entry: the breaker, the in-hook count WaitForHooksQuiet reads,
        // the fault handler, and a guard against re-entry on one thread (VBA run
        // by a COM call made from inside a hook).
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
    // Did somebody catch it? A statement running again means the error was dealt with, by the
    // frame it ran in. Called first in OnStatementBody, before the fast path returns.
    //
    // errActive is set when the thrower closes, so the first statement that observes it is the
    // handler's second, by which point the unwind's closes are done and the top of the stack is
    // the frame that resumed. Only a frame that predates the throw can be the handler: a
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

    // Does an activation start here? vbaboundary.h owns the answer. Asked before the fast path:
    // a UDF that raises never reaches its epilogue, so its next call arrives with the same
    // trailer at the same rsp, which is the fast path's exact condition. Unavailable is counted
    // apart from No.
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

    // Same procedure, DIFFERENT stack pointer. Counting only -- nothing
    // downstream reads these; they exist so the fast path can be SEEN working.
    void NoteStackPointerMove(ThreadState* s, std::uint64_t trailer, std::uint64_t dispatchSp)
    {
        if (!(trailer == s->current && s->currentSp != 0)) return;
        if (dispatchSp < s->currentSp)
            Bump(g_totals.sameTrailerDeeper);
        else if (dispatchSp > s->currentSp)
            Bump(g_totals.sameTrailerShallower);
    }

    // Closes what has returned before anything is opened: the stack grows down, so a frame
    // whose stack has been released has a smaller sp than the current one.
    //
    // Beyond the shadow stack's cap nothing was stored, so the only evidence a capped
    // activation ended is a shallower rsp. That undercounts a multi-level unwind, which is why
    // `deepestSeen` is published as a floor.
    void CloseFramesThatReturned(ThreadState* s, std::uint64_t dispatchSp, std::uint64_t now)
    {
        while (s->depth > 0 && s->stack[s->depth - 1].sp < dispatchSp)
            CloseFrame(s, now);

        if (s->overflowDepth > 0 && s->currentSp != 0 && dispatchSp > s->currentSp)
            --s->overflowDepth;
    }

    // Did Excel start this frame, rather than VBA calling it? Excel enters VBA to compute a
    // worksheet function, and xlfCaller names the cell. A VBA call within that calc keeps the
    // same cell, so a caller cell that differs from the frame beneath, or a cell where the
    // frame beneath has none, is a fresh worksheet-function entry: the one activation whose
    // unhandled error Excel turns into #VALUE! instead of passing to a VBA caller.
    //
    // A sheet event has no calling cell, so a user-triggered event's unhandled error reads
    // `threw` rather than `unhandled`.
    bool IsExcelTheCaller(ThreadState* s, bool callerIsCell, std::uint64_t callerHash)
    {
        if (!callerIsCell) return false;
        return s->depth < 2 || s->stack[s->depth - 2].callerHash != callerHash;
    }

    // Is a frame on this thread waiting inside DoEvents, with nothing of its own since?
    // Excel runs timer macros and events from in there, so such a procedure is not its callee.
    // A mark left behind -- End pressed inside DoEvents runs no epilogue -- is dropped here,
    // because the frame it was made at is gone.
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

    // Only a prologue opens a frame. Opening on the stack pointer is wrong, because `GoSub`
    // moves rsp within one activation; rsp still closes frames.
    //
    // The one case with no prologue is a procedure already running when tracing armed. It is
    // opened anyway and counted in `lateOpen`.
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

    // Who called this, once per activation: xlfCaller is a call into Excel from a traced
    // thread. Always asked, since the VBA tracer needs the cell to place an error that escapes
    // into it. Four facts counted apart: not asked, a cell, a caller that is not a cell, and
    // Excel declining.
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

    // OPEN AND DESCRIBE A NEW ACTIVATION. The ORDER inside is load-bearing:
    // stale-close before push, and args captured at the FIRST statement before
    // the body can overwrite a ByRef. On overflow it returns having emitted the
    // single depth-capped marker; either way the caller updates current/currentSp.
    void OpenFrame(ThreadState* s, std::uint64_t trailer, std::uint64_t savedRegs,
                   std::uint64_t dispatchSp, bool atEntry, bool lateOpen,
                   std::uint64_t now)
    {
        if (lateOpen)
            Bump(g_totals.ipLateOpen);

        // A frame still open at this exact rsp cannot still be running: this
        // statement has taken its place. When it is the SAME procedure it is the
        // frame a raising UDF left behind.
        if (s->depth > 0 && s->stack[s->depth - 1].sp == dispatchSp)
        {
            if (atEntry && s->stack[s->depth - 1].trailer == trailer)
                Bump(g_totals.ipStaleClosed);
            CloseFrame(s, now);
        }

        if (s->depth >= kMaxDepth)
        {
            Bump(g_totals.overflows);

            // No frame, but the DEPTH is still knowable and still matters:
            // without it the totals report the cap for a deeper recursion, which
            // is a wrong answer rather than a truncated one.
            ++s->overflowDepth;
            RaiseMax32(&g_totals.deepestSeen,
                       static_cast<std::uint32_t>(s->depth + s->overflowDepth));

            // SAY IT IN THE TRACE, NOT ONLY IN THE TOTALS: the CSV is what a
            // user reads, and there a call tree that stops at the cap looks
            // exactly like a recursion that ended there. One marker row per
            // arming session is enough to stop the tree lying about its end.
            if (InterlockedCompareExchange(&g_cappedNoted, 1, 0) == 0)
            {
                Frame m{};
                m.trailer = trailer;
                m.span    = 0;
                m.parent  = (s->depth > 0) ? s->stack[s->depth - 1].span : 0;
                EmitRow("depth-capped", m, Now(), 0, s->depth, nullptr);
            }
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
        // -1 because the statement that OPENED this frame is already counted
        // above; otherwise the callee's first statement is billed to its caller.
        f.stmtsAtEntry = s->statements - 1;
        f.sp           = dispatchSp;
        f.span         = emit::csv::NextSpan();
        f.parent       = parent;
        f.shownDepth   = shownDepth;
        f.outcome      = kOutReturned;
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
        // R14 is the VBA frame base. Arguments are captured here, at the first statement,
        // before the body can overwrite a ByRef one. ReadReg's answer is checked, so a broken
        // assembly contract is not mistaken for a frame with no base.
        std::uint64_t r14 = 0;
        const bool haveR14 = ReadReg(savedRegs, kReg_r14, r14);
        if (!haveR14)
            Bump(g_totals.regReadFailures);

        ArgCapture args;
        args.text = &t_argRender;                   // the per-thread render buffer
        // ARGS latched at arm: when off the column is empty and the
        // decline is counted, so it never reads as "no arguments".
        if (haveR14 && InterlockedCompareExchange(&g_capArgs, 0, 0) != 0)
        {
            CaptureArgs(trailer, r14, args);
            // Remembered so the EXIT can tell whether anything moved. See
            // the Frame fields: the hash, not the text.
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

        // THE ORDER OF THESE STEPS IS LOAD-BEARING. They read top to bottom as
        // the question this hook answers: was an error caught, where are we, is
        // this the same frame, what has returned, does a frame open, and what
        // goes in the row.
        MarkHandlerIfErrorResumed(s);

        std::uint64_t trailer = 0;
        // Zero is no procedure: nothing to name, count or open a frame for.
        if (!ReadTrailer(dispatchSp, trailer) || trailer == 0)
        {
            Bump(g_totals.faults);
            return;
        }

        const bool atEntry = IsActivationStart(trailer, savedRegs);

        // THE FAST PATH. Same procedure, same interpreter stack pointer: this
        // is the next statement of the frame we are already in, and there is
        // nothing to do. Everything below happens only on a real frame change.
        if (!atEntry && trailer == s->current && dispatchSp == s->currentSp)
        {
            Bump(g_totals.sameTrailerSameSp);
            return;
        }

        NoteStackPointerMove(s, trailer, dispatchSp);

        const std::uint64_t now = Now();
        Bump(g_totals.transitions);

        // The stack pointer is part of the frame identity: the trailer identifies a procedure,
        // not an activation, so recursion would be invisible. rsp at the dispatch falls one
        // step per VBA call and is stable within a frame, so (trailer, sp) identifies an
        // activation.

        CloseFramesThatReturned(s, dispatchSp, now);
        if (atEntry) CloseFramesThatLostTheirStack(s, now);

        bool lateOpen = false;
        if (ShouldOpenFrame(s, trailer, atEntry, dispatchSp, lateOpen))
            OpenFrame(s, trailer, savedRegs, dispatchSp, atEntry, lateOpen, now);

        s->current   = trailer;
        s->currentSp = dispatchSp;
    }

    // THE EXIT OPCODE CLOSES ITS OWN FRAME -- the one signal that fires exactly
    // once per activation. The stack pointer stays the BACKSTOP for frames
    // abandoned without an exit, since an error unwind and `End` fire none at
    // all. Double-closing is prevented by requiring a match on BOTH trailer and
    // sp; anything else is counted, not acted on.
    void OnExitBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
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
            CloseFrame(s, Now(), r14, exitOp, "exit");
            // Forget the fast-path cache, or the NEXT activation of the same
            // procedure at the same rsp matches it and never opens a frame.
            s->current   = 0;
            s->currentSp = 0;
            Bump(g_totals.exitClosed);
        }
        else
        {
            Bump(g_totals.exitNoMatch);
        }
    }
        // `End` tears the whole VBA session down and fires no exit opcode for any frame it
        // kills. The backstop alone cannot close them when the next chain starts deeper: rsp
        // cannot tell "nested" from "the last chain was abandoned".
        //
        // This thread only, because the shadow stack is TLS. Ticks are a measurement here: the
        // opcode fires when the session dies.
        void OnEndBody()
        {
            ThreadState* s = State();
            if (!s) return;
            const std::uint64_t now = Now();
            while (s->depth > 0)
            {
                s->stack[s->depth - 1].outcome = kOutAbandoned;
                CloseFrame(s, now, 0, 0, "end");
            }
            // The chain is gone, so nothing about it may survive to be matched
            // against the next one -- including the fast path's (trailer, sp) and
            // the depth we could not store frames for.
            s->overflowDepth = 0;
            s->current = 0; s->currentSp = 0;
            s->errActive = false; s->errSpan = 0;
        }

    }   // anonymous namespace

    // -------------------------------------------------------------------

    // Called by the rtcDoEvents detour, on the thread that is entering or leaving it. Nesting
    // is kept because a macro run inside DoEvents may call DoEvents itself; past the small
    // stack the outermost marks are kept and the innermost dropped, which only costs the
    // parenting of a chain nested that deep.
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

    void FlushOpenFrames()
    {
        // t_state is the storage itself, not a pointer to it, so a thread that
        // never traced has generation 0 and depth 0 and this is a no-op for it.
        ThreadState* s = &t_state;
        const std::uint64_t now = Now();
        // A frame this thread is still inside sits above the current stack with its
        // trailer in place. A leftover from an earlier unwind has lost that stack.
        const std::uint64_t here = reinterpret_cast<std::uint64_t>(_AddressOfReturnAddress());
        bool running = false;
        while (s->depth > 0)
        {
            const Frame& top = s->stack[s->depth - 1];
            // Still running only if it sits above our own stack AND its trailer is in
            // place; a leftover from an earlier unwind fails the second test. A frame
            // beneath a running one is running too, so one chain is never split.
            running = running || (top.sp > here && FrameStillLive(top));
            CloseFrame(s, now, 0, 0, "flush", running);
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
        // A name being republished right now could tear. The trailer is
        // always true, so the caller gets that instead.
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
        // A new session gets a new "the stack ran out" marker row, or a run
        // that capped after an earlier one would emit nothing and read clean.
        InterlockedExchange(&g_cappedNoted, 0);
        ResetIdentityCounts();
        ResetArgCounts();
        // Otherwise the p-code declines accumulate across every arm/disarm for
        // the life of the process -- including `Desynced`, the one that means a
        // type may have been attributed to the WRONG argument.
        ResetPcodeCounts();
        // Arming is a deliberate human act, so it is the right place to give
        // the tracer another go -- from a fault count of zero, not from where it
        // left off.
        g_breaker.faults = 0;
        InterlockedExchange(&g_breaker.tripped, 0);
        // Per ARMING SESSION, like everything else here: left to accumulate
        // they report opcodes from a previous arm as if this one had seen them.
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
