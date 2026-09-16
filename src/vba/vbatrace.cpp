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
        // An unhandled error left VBA through a frame Excel called for a cell (D92).
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
            // HOW THIS ACTIVATION ENDED.
            //
            //   returned  it finished normally
            //   threw     the error was raised IN this frame
            //   unwound   the error passed THROUGH it -- it ran nothing after the
            //             raise, so it did not handle it
            //   handled   it ran again after the raise, so it caught it
            //
            // A chain reads threw -> unwound -> unwound -> handled from the throwing
            // frame outwards: where was it thrown, and who caught it.
            std::uint8_t  outcome;

            // DID THE RAISE OPCODE FIRE IN THIS FRAME? The outcome is decided
            // at CLOSE, because at the raise a real Err.Raise and ordinary
            // object-model VBA (a cell write, For Each, OnTime) are identical --
            // same opcode, same registers. How the frame LEFT tells them apart.
            // Deferring also stops a benign raise's provisional error state
            // suppressing a genuine throw in a nested frame.
            bool          raised;

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

            // No caller gate: the exit opcode says whether a result exists (a
            // Sub leaves through slot 635) and of what kind, so the return is
            // read at every depth and whatever the caller was.
        };

        // A PLAIN AGGREGATE with no member initialisers, so it is trivially
        // constructible and can live in thread-local storage directly: the
        // loader zero-fills it and no dynamic initialiser runs on a VBA thread.
        // `generation` starting at zero is load-bearing -- g_generation starts
        // at 1, so the ordinary check below initialises the block on first use
        // with no "first time?" flag.
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

            // ---- AN ERROR IN FLIGHT ON THIS THREAD -----------------------
            //
            // THE RULE: an error was raised, and a
            // frame that ran NOTHING SINCE did not handle it -- it unwound. A
            // statement running again in a frame that PREDATES the raise means
            // that frame caught it.
            //
            // errActive is set at the THROWER'S CLOSE, not at the raise, so the
            // first statement at which MarkHandlerIfErrorResumed observes it is
            // already the settled one -- and so a benign raise's provisional
            // error state cannot dedupe, and thereby suppress, a genuine throw
            // in a nested frame. Setting it eagerly at the raise passes every
            // single-raise case and loses that one.
            bool          errActive;
            // THE SPAN OF THE FRAME THAT THREW. Spans only ever increase, so a frame
            // with a LOWER span existed before the raise and can be the handler; a higher
            // one opened during the unwind and cannot.
            std::uint64_t errSpan;
            // THE RAISING FRAME IS USUALLY NOT OPEN YET: a procedure opens at
            // its FIRST statement, so a UDF beginning with Err.Raise raises one
            // opcode before it exists, and marking the top of the stack would
            // give the error to its CALLER. Held here, applied when it opens.
            bool          errPending;
        };

        // TLS, because hooks run reentrantly across threads. The STORAGE is
        // thread-local, not a thread-local pointer to the heap.
        __declspec(thread) ThreadState t_state;

        // The rendered ARGUMENT LIST buffer: heap, lazily allocated on
        // each thread's first hook call, so only threads that run a VBA hook pay
        // for it and nothing is resized per row. On the heap rather than the
        // stack so ArgCapture stays tiny on the deeply nested interpreter stack.
        // Leaked at thread exit, bounded to one block per hooking thread.
        constexpr int kArgRenderMax = 64 * 1024;
        __declspec(thread) char* t_argRender = nullptr;
        char* ArgRenderBuf()
        {
            if (!t_argRender) t_argRender = static_cast<char*>(malloc(kArgRenderMax));
            return t_argRender;   // may be null -- CaptureArgs then declines cleanly
        }

        // The return column's buffer, on the same terms.
        constexpr int kRetRenderMax = 16384;
        __declspec(thread) char* t_retRender = nullptr;
        char* RetRenderBuf()
        {
            if (!t_retRender) t_retRender = static_cast<char*>(malloc(kRetRenderMax));
            return t_retRender;
        }

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
        // ARGUMENT CAPTURE CAN BE TURNED OFF, as a control. It is the only part
        // of the hook that FOLLOWS POINTERS out of the frame -- a BSTR, an
        // object, a SAFEARRAY, any ByRef -- and the only part that walks a
        // procedure's bytecode, so it has the largest fault surface and is the
        // first thing to remove when asking whether the tracer is what is
        // killing Excel. Frames, names, timing and the call tree are unaffected,
        // so a run with this off is still a trace.
        //
        // Latched at arm. A disabled decode leaves the column empty and is
        // COUNTED, never silently blank.
        volatile LONG g_capArgs = 1;

        // DOES THE SIGNATURE CONTAIN A ByRef PARAMETER? ByRef is VBA's DEFAULT,
        // so `Sub Calc(result As Double)` filling in `result` is ordinary code
        // whose whole effect a trace of the entry values would miss.
        //
        // The marker is the "&" the type table already uses, and the same test
        // the argument decoder uses to decide to dereference -- so a signature
        // this says yes to is exactly one whose values it can follow.
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

        // The SAME trace file as the XLL side, and one SEQUENCE across both,
        // because the product is one causal timeline rather than two.
        //
        // `seq` and `qpc` are not the same thing: a row is STAMPED when its
        // event happened and NUMBERED when it is written, and the work between
        // (arguments, names, the caller) is why they can differ. seq orders the
        // file; qpc orders the events.
        //
        // Everything here runs on a frame change, never per statement, so the
        // synchronous write is off the hot path even though the file is not.
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
            // WHERE THIS SAT IN THE CHAIN, on every VBA row, entry and exit
            // alike. `parent` is a span, so a missing row shows as a reference
            // to a span that is not in the file -- a question, where rebuilding
            // the tree from `depth` and row order alone would silently attach
            // everything after a gap to the wrong caller.
            _snprintf_s(parentb, _TRUNCATE, "%llu",
                        static_cast<unsigned long long>(f.parent));
            _snprintf_s(depthb, _TRUNCATE, "%d", depth);
            // On the EXIT row only -- an entry row cannot know yet. `returned`
            // is written explicitly rather than left off, because this column
            // exists precisely so an error unwind stops looking like a clean
            // return.
            // WHAT CLOSED THE FRAME, because it decides whether `ticks` is a
            // MEASUREMENT or an UPPER BOUND. Only the exit opcode fires at the
            // moment an activation ends; the stack-pointer backstop fires at the
            // NEXT statement and the flush at disarm.
            //
            // It matters most where a user is most likely to be looking: a fully
            // unhandled unwind fires ZERO exit opcodes, so every frame in that
            // chain is closed by the backstop and would otherwise read as a slow,
            // successful call.
            //
            // The number is kept: a true upper bound is still a fact.
            // TWO COLUMNS, NOT A SENTENCE. An entry row has no duration, so
            // both stay empty there rather than carrying a nothing.
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
                argsText = args->text;
                if (!argValuesOnly)
                {
                    _snprintf_s(argcb, _TRUNCATE, "%d", args->params);
                    sigText = args->signature;   // "(Long,String)", or empty
                }
            }

            // `proc` is always the trailer -- the identity the tracer actually
            // used, which tells two same-named procedures apart. `module` is
            // EMPTY when unresolvable, never a guess; `function` falls back to
            // the trailer in hex, which obviously looks like an address.
            //
            // The caller's STACK-LOCAL resolution is preferred over the shared
            // Proc entry, which is rewritten on every push and exists for the
            // end-of-session Report.
            ResolvedName shared{};
            if (p && !(nm && nm->qualModule[0] && nm->function[0])) SnapshotName(p, shared);
            const char* mod  = (nm && nm->qualModule[0]) ? nm->qualModule : shared.qualModule;
            const char* func = (nm && nm->function[0])   ? nm->function
                             : (shared.function[0] ? shared.function : trailerb);

            // Only an entry row carries this: an exit is the same activation as
            // the entry it pairs with. `cell` and `sheet` stay EMPTY for
            // anything that is not a real cell on a real sheet -- a macro on a
            // button has an object name, and `caller` is where that goes.
            // Empty rather than wrong (core/caller.h).
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
                // The PROJECT name is deliberately left out: objTable+0x90
                // yields "ThisWorkbook" on this build -- a document module, not
                // the project -- so printing it would put a confidently
                // incorrect name on every row. Workbook and module are verified
                // and identify a procedure anyway.
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

        // ASKING EXCEL FROM A THREAD EXCEL IS NOT CALLING US ON. The C API is
        // documented for use inside an XLL callback; here we are on Excel's
        // calculating thread but inside VBE7's interpreter, which Excel reached
        // through COM. That xlfCaller answers at all was MEASURED, not assumed
        // -- it does, because it is Excel's per-thread state and the calculating
        // thread is where that state lives.
        //
        // Guarded anyway, in its own leaf frame: a fault here would otherwise be
        // attributed to the tracer's own hook and trip the circuit breaker,
        // standing down VBA tracing entirely over a question we could have
        // declined to answer.

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
                s->errPending = false; s->errSpan = 0;
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

        // ---- DECIDE THE OUTCOME AS THE FRAME CLOSES -------------------------
        //
        // The raise opcode (497) fires for BOTH a real Err.Raise and ordinary
        // object-model VBA, and nothing AT the raise separates them. The frame's
        // FATE does, and is known only here: one that raised and ran its own
        // epilogue resolved it in place (benign), one that raised and unwound
        // with no epilogue really threw -- and the frames OUTER than it unwind
        // through until one resumes.
        //
        // `ranEpilogue` is (exitOp != 0): a real unhandled throw fires no exit
        // opcode, closing through the stack-pointer backstop instead.
        void DecideOutcomeAtClose(ThreadState* s, Frame& f, bool ranEpilogue, bool stillRunning)
        {
            // Already resolved as the catcher -- unless it then raised an error that left it,
            // which makes it that error's thrower.
            if (f.outcome == kOutHandled && !(f.raised && !ranEpilogue)) return;
            // `End` already settled this one, and no epilogue ran anywhere, so
            // every rule below would read it as an unwind it was not part of.
            if (f.outcome == kOutAbandoned) return;
            // Disarm closed it while it was still running: a raise in it is not a throw.
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

            if (f.raised && ranEpilogue)
            {
                // Raised, then exited normally: a benign object-model raise
                // VBE7 resolved in place, or a same-frame handler. Not a throw.
                // A same-frame On Error Resume Next of a REAL error also lands
                // here as `returned` -- the rarer case, and the less misleading
                // of the two to lose.
                f.outcome = kOutReturned;
                Bump(g_totals.raisesBenign);
                return;
            }

            // A DEEPER error already unwinding (span greater than this frame's)
            // is the fatal one; this frame's own raise, if any, is beside it.
            const bool deeperError = s->errActive && s->errSpan > f.span;

            if (f.raised && !ranEpilogue && !deeperError)
            {
                // Raised and unwound with no epilogue: THIS frame threw and owns
                // the episode, and the frames outside it now unwind through.
                f.outcome     = kOutThrew;
                Bump(g_totals.threw);
                s->errActive  = true;
                s->errSpan    = f.span;
                return;
            }

            // NOT THE THROWER. Only a frame OUTER than the raise is unwound
            // THROUGH; one the raiser had called INTO (higher span) is a
            // finished nested call, not a casualty.
            if (s->errActive && f.span < s->errSpan && !ranEpilogue &&
                f.outcome == kOutReturned)
            {
                f.outcome = kOutUnwound;
                Bump(g_totals.unwound);
            }
        }

        // ---- WHAT KIND OF RESULT THIS PROCEDURE LEFT BEHIND ------------------
        //
        // The exit opcode says, for most procedures (vbaretdecode.h). Class and
        // form Functions all leave through 1664 whatever they return, so for
        // them the instruction that WROTE the result is asked instead.
        //
        // SCANNING AND TYPING ARE SEPARATE JOBS: vbapcode.cpp finds every candidate
        // store to the result slot and knows nothing about what the opcodes
        // mean; vbaretdecode.cpp owns that.
        //
        // THE OFFSET IDENTIFIES A VARIANT, not the opcode: a Variant result
        // lives at [R14-0x18] and nothing else does, and it is stored through
        // DIFFERENT opcodes depending on what it holds. The slot is the
        // invariant.
        //
        // TWO KNOWN STORES OF DIFFERENT TYPES CANNOT BOTH BE THE RESULT, and
        // nothing here says which is the false positive, so the type is refused.
        //
        // Pure: a function of the bytecode and the exit opcode.
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

        // ---- ByRef ARGUMENTS THE PROCEDURE CHANGED ---------------------------
        //
        // The entry captured the arguments at the FIRST statement, before the
        // body could touch them; this is the "after". A ByRef slot points at the
        // CALLER's variable, and the caller's frame is still live at the exit
        // opcode, so the pointer still resolves.
        //
        // BYVAL IS DELIBERATELY NOT REPORTED: a ByVal slot holds a copy the
        // caller never sees, so printing it as an "after" value would assert an
        // effect that does not exist. `hasByRef` is the gate and the cost bound.
        //
        // Returns `&out` when the arguments MOVED; the four counters keep
        // "nothing changed" apart from "never looked".
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
            if (out.byRefHash != f.argsHash)
            {
                Bump(g_totals.byrefChanged);
                return &out;
            }
            Bump(g_totals.byrefSame);
            return nullptr;
        }

        // ---- HOW THE RESULT READ WENT, in the totals ---------------------------
        //
        // Read, declined, or never attempted because the exit opcode has no
        // mapping -- a gap in the TABLE, not a refusal by the decoder. The store
        // candidates are recorded for that case, so it can be closed from
        // evidence.
        void CountReturnOutcome(std::uint64_t trailer, std::uint64_t r14, RetKind kind,
                                std::uint16_t exitOp, bool haveRet)
        {
            if (haveRet)
                Bump(g_totals.returnsRead);
            else if (r14 != 0 && kind != RetKind::None && kind != RetKind::Unknown)
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

        // ---- CLOSE THE FRAME ON TOP ---------------------------------------------
        //
        // `r14` is the frame base of the activation that is ENDING, and comes
        // ONLY from the exit-opcode path. The other close paths run when a frame
        // is DISCOVERED to have already returned, where R14 belongs to a
        // different activation entirely and reading a return value would
        // attribute one procedure's memory to another.
        //
        // `exitOp` is the typed exit that names the return kind (vbaretdecode.h),
        // or 0 from the paths that close without having seen it.
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
            // THE RESULT, at any depth. The exit opcode says what kind it is
            // and whether one exists at all; vbaretdecode.cpp reads it from where
            // that kind lives and refuses what it cannot vouch for. Slot 634
            // alone needs the store opcode to split LongLong from an array, and
            // only then is the body scanned.
            //
            // The same size as the argument column's buffer, so an array renders the
            // same in both. Per-thread heap, not the stack VBA may be about to run out of.
            char* const retText = RetRenderBuf();
            const char* retType = "";
            bool haveRet = false;
            std::uint16_t storeOp = 0;
            const RetKind kind = DecideReturnKind(f.trailer, exitOp, storeOp);
            const bool readable = (r14 != 0 && kind != RetKind::None && kind != RetKind::Unknown);
            if (retText && readable && InterlockedCompareExchange(&g_capRet, 0, 0) != 0)
            {
                retText[0] = 0;
                haveRet = DescribeReturnKind(r14, kind, storeOp, retText, kRetRenderMax, &retType);
            }
            // The "after" for the entry's "before": see ReadByRefChanges. Reuses
            // the same per-thread render buffer -- the entry's text is already
            // written and only its hash is kept, so nothing live is overwritten.
            ArgCapture outArgs;
            outArgs.text = ArgRenderBuf();
            outArgs.textCap = outArgs.text ? kArgRenderMax : 0;
            const ArgCapture* outArgsPtr = ReadByRefChanges(f, r14, outArgs);

            if (!g_topLevelOnly || s->depth == 1)
                EmitRow("exit", f, nowTicks, nowTicks - f.startTicks, s->depth, p,
                        outArgsPtr, nullptr, nullptr,
                        haveRet ? retText : "",
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

        // A STACK OVERFLOW IS NOT AN ORDINARY FAULT. The hook runs at exactly
        // the depths where the least stack remains, because VBA recursion runs
        // until VBE7 itself gives up, and catching one without restoring the
        // guard page leaves the thread primed to die later in unrelated code.
        //
        // So the filter remembers the code and the handler puts the guard page
        // back (_resetstkoflw is documented to be called from exactly here) and
        // opens the breaker AT ONCE rather than after eight faults.
        __declspec(thread) unsigned t_lastExceptionCode = 0;


        // EVERY code is handled here, and that is LOAD-BEARING, not laziness.
        // vbathunk.asm is declared PROC FRAME with .endprolog BEFORE its fifteen
        // pushes, so its unwind info describes an EMPTY prologue -- safe only
        // while nothing ever unwinds through it. Return
        // EXCEPTION_CONTINUE_SEARCH for any code and the unwinder computes RSP
        // as though those ~264 bytes were never pushed, takes a garbage return
        // address, and transfers execution to whatever it points at: the exact
        // signature of the one unexplained crash this project has seen.
        //
        // If a code ever needs to be declined, fix the thunk's unwind info FIRST
        // (.pushreg per push and .setframe rbx). The two are one mechanism, and
        // nothing else ties them together.
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
        void OnRaiseBody(std::uint64_t dispatchSp, std::uint64_t savedRegs);
        void OnEndBody();
    }

    // THE HOOK BODIES ARE SEH-WRAPPED. Guarding the individual reads of
    // interpreter memory is not enough: a fault in the tracer's OWN code -- name
    // resolution, argument decoding, the p-code walk, the emitter -- would
    // propagate straight into the VBA interpreter on a live calc thread.
    //
    // Split into wrapper + body because __try may not share a frame with objects
    // that need unwinding, and the bodies have several.
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

    extern "C" void XRayVbaOnRaise(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnRaiseBody, dispatchSp, savedRegs);
    }

    // The End opcode has no frame to read; its arguments are unused.
    extern "C" void XRayVbaOnEnd(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        RunHook(OnEndAdapter, dispatchSp, savedRegs);
    }

    namespace
    {
    // AN ERROR WAS RAISED HERE.
    //
    // The opcode fires FOUR TIMES for one Err.Raise, so the first wins and the
    // rest are counted and discarded. Dedupe is PER FRAME: a second raise in the
    // same activation, after the first resolved, is a real second error.
    void OnRaiseBody(std::uint64_t dispatchSp, std::uint64_t savedRegs)
    {
        (void)savedRegs;
        (void)dispatchSp;
        ThreadState* s = State();
        if (!s) return;

        // MARK THAT A RAISE FIRED HERE -- DECIDE NOTHING YET. Whether this is a
        // real throw or a benign object-model raise is unknowable at the raise
        // itself, so the outcome waits for DecideOutcomeAtClose. A per-frame flag
        // is idempotent, so the extra fires are counted as deduped rather than
        // acted on twice.
        if (s->depth <= 0)
        {
            // Nothing open to attribute it to: the raising procedure has not
            // reached its first statement. Held for the next frame to open --
            // marking the top of the stack would hand the error to the CALLER.
            if (s->errPending)
            {
                Bump(g_totals.raisesDeduped);
                return;
            }
            s->errPending = true;
            Bump(g_totals.raises);
            Bump(g_totals.errNoFrame);
            return;
        }

        Frame& f = s->stack[s->depth - 1];
        if (f.raised)
        {
            Bump(g_totals.raisesDeduped);
            return;
        }
        f.raised = true;
        Bump(g_totals.raises);
    }

    // ---- DID SOMEBODY CATCH IT? -------------------------------------------
    //
    // A statement running again means the error was dealt with, and the frame it
    // ran in is the one that dealt with it.
    //
    // CALLED FIRST IN OnStatementBody, before the fast path returns -- otherwise
    // most statements never reach it.
    //
    // THE FIRST STATEMENT THAT SEES THE ERROR is already the settled one:
    // errActive is set when the THROWER closes, and that close happens inside
    // the processing of the handler's first statement, AFTER this check ran for
    // it. So the first statement at which this observes errActive is the
    // handler's second, by which point the unwind's closes are done and the top
    // of the stack is the frame that resumed.
    //
    // ONLY A FRAME THAT PREDATES THE RAISE CAN BE THE HANDLER: a destructor
    // opened during the unwind runs statements of its own but has a HIGHER span.
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

    // DOES AN ACTIVATION START HERE? vbaboundary.h owns the answer and its
    // evidence; this file only acts on it.
    //
    // ASKED BEFORE THE FAST PATH, which is the whole point: the defect it fixes
    // is a second activation with the SAME trailer at the SAME rsp, which is the
    // fast path's exact condition. A UDF that raises never reaches its epilogue,
    // so its frame stays open and the next call is waved through as the next
    // statement of the frame we are already in.
    //
    // Unavailable is counted apart from No, so a boundary that quietly fell back
    // to the rsp heuristic does not print the same number as one that was never
    // wrong. Hence a bool return and three counted outcomes.
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

    // WHAT HAS RETURNED, closed before anything is opened: the stack grows
    // down, so anything whose stack has been released has a SMALLER sp than we
    // are now at.
    //
    // COMING BACK UP PAST THE CAP is the second half of the same job. Beyond the
    // shadow stack nothing was stored, so there is no frame to match and the
    // only evidence a capped activation ended is a SHALLOWER rsp than it was
    // last seen at. That undercounts a multi-level unwind, which is why
    // `deepestSeen` is published as a floor.
    void CloseFramesThatReturned(ThreadState* s, std::uint64_t dispatchSp, std::uint64_t now)
    {
        while (s->depth > 0 && s->stack[s->depth - 1].sp < dispatchSp)
            CloseFrame(s, now);

        if (s->overflowDepth > 0 && s->currentSp != 0 && dispatchSp > s->currentSp)
            --s->overflowDepth;
    }

    // IS EXCEL A LIVE CALLER between this VBA frame and its parent? Walk the REAL
    // DID EXCEL START THIS FRAME, rather than VBA calling it? Excel enters VBA to
    // compute a worksheet function, and xlfCaller names the cell it is computing. A
    // VBA call within that calc keeps the same cell, so a caller cell that DIFFERS
    // from the frame beneath -- or a cell where the frame beneath has none -- is a
    // fresh worksheet-function entry, the one kind of activation whose unhandled error
    // Excel turns into that cell's #VALUE! instead of passing to a VBA caller. Errors
    // through Application.Run or an object-model call are NOT cells, so they propagate.
    //
    // A sheet event has no calling cell, so a user-triggered event's unhandled error
    // reads `threw` rather than `unhandled` (D92).
    bool IsExcelTheCaller(ThreadState* s, bool callerIsCell, std::uint64_t callerHash)
    {
        if (!callerIsCell) return false;
        return s->depth < 2 || s->stack[s->depth - 2].callerHash != callerHash;
    }

    // A FRAME THAT LOST ITS STACK. An unhandled error runs no epilogue, and Excel may
    // call the next cell's function at the same depth or deeper, so the backstop never
    // fires. A running frame keeps its trailer on the stack; a dead one does not.
    // Is this frame still executing? The interpreter keeps a frame's trailer at the
    // dispatch rsp we recorded for it, so a returned frame's slot holds something else.
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

    // WHO OPENS A FRAME: only a prologue. Every genuine activation begins with
    // one (459 activations, 28 shapes, three entry routes), so the stack pointer
    // is not needed to open anything -- and opening on it is wrong,
    // because `GoSub` moves rsp WITHIN one activation, which reads as a nested call. rsp still CLOSES frames, which it does correctly.
    //
    // The one case with no prologue to go on is a procedure ALREADY RUNNING when
    // tracing armed. Its trailer is on nobody's stack, so it is opened anyway
    // and counted as a guess, which is what `lateOpen` reports back.
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

    // WHO CALLED THIS -- once per activation, the only granularity at which it
    // is affordable: xlfCaller is a call INTO Excel from a traced thread, which
    // is a handful per calc here and would be on the hot path per statement.
    //
    // FOUR different facts, counted apart: not asked, a cell, a caller that is
    // not a cell (a button, a toolbar, an event -- answers, not failures), and
    // Excel declining to answer.
    //
    // ALWAYS ASKED. The calling cell is part of every row and the VBA tracer needs
    // it to place an error that escapes into a cell (D92); the lookup is a per-entry
    // Excel12 round-trip inside the span being timed.
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

        const std::uint64_t parent = (s->depth > 0) ? s->stack[s->depth - 1].span : 0;
        Frame& f = s->stack[s->depth++];
        f.trailer      = trailer;
        f.startTicks   = now;
        // -1 because the statement that OPENED this frame is already counted
        // above; otherwise the callee's first statement is billed to its caller.
        f.stmtsAtEntry = s->statements - 1;
        f.sp           = dispatchSp;
        f.span         = emit::csv::NextSpan();
        f.parent       = parent;
        f.outcome      = kOutReturned;
        f.raised       = false;
        f.hasByRef     = false;
        f.argsHash     = 0;
        f.fromExcel    = false;
        f.callerHash   = 0;

        // A RAISE THAT ARRIVED BEFORE ITS FRAME EXISTED: Err.Raise on the first
        // line fires one opcode before the procedure opens, so the mark was held
        // rather than given to the caller. This is that frame, and like any
        // raise its fate is decided at close.
        if (s->errPending)
        {
            f.raised      = true;
            s->errPending = false;
            Bump(g_totals.errLateOpen);
        }
        Bump(g_totals.framesOpened);

        ResolvedName nm{};
        Proc* p = FindProc(trailer);
        if (p)
        {
            ResolveInto(p, trailer, &nm);   // fresh every call -- never cached
            Bump(p->calls);
            RaiseMax64(&p->maxDepth, static_cast<std::uint64_t>(s->depth));
        }
        // R14 is the VBA frame base. Arguments are captured HERE, at the first
        // statement, because that is when they are in place and before the body
        // can overwrite them -- a ByRef argument assigned in the body would
        // otherwise be reported as whatever it became.
        //
        // ReadReg's answer is checked: discarding it merges "the assembly
        // contract broke" with "this frame legitimately had no frame base".
        std::uint64_t r14 = 0;
        const bool haveR14 = ReadReg(savedRegs, kReg_r14, r14);
        if (!haveR14)
            Bump(g_totals.regReadFailures);

        ArgCapture args;
        args.text = ArgRenderBuf();                 // the per-thread render buffer
        args.textCap = args.text ? kArgRenderMax : 0;
        // ARGS latched at arm: when off the column is empty and the
        // decline is counted, so it never reads as "no arguments".
        if (haveR14 && InterlockedCompareExchange(&g_capArgs, 0, 0) != 0)
        {
            CaptureArgs(trailer, r14, args);
            // Remembered so the EXIT can tell whether anything moved. See
            // the Frame fields: the hash, not the text.
            if (args.ok)
            {
                // EITHER TEST IS ENOUGH and they catch different things: the
                // signature names a ByRef type the p-code declared, while
                // `viaPointer` catches a slot dereferenced to read its value
                // even though nothing declared it -- a write-only `String`, the
                // one shape that recovers no type from a load OR a store.
                f.hasByRef = SignatureHasByRef(args.signature) || args.viaPointer;
                f.argsHash = args.byRefHash;
            }
        }

        core::Caller who;      // NOT {} -- IdentifyCaller fills it on every path
        IdentifyCaller(who);
        f.callerHash = who.isCell ? core::Fnv1aText(who.desc) : 0;
        f.fromExcel  = IsExcelTheCaller(s, who.isCell, f.callerHash);

        if (!g_topLevelOnly || s->depth == 1)
            EmitRow("entry", f, now, 0, s->depth, p, &args, &nm, &who);

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

        // WHY THE STACK POINTER IS PART OF THE FRAME IDENTITY: the trailer
        // alone identifies a PROCEDURE, not an activation of it, so a procedure
        // calling itself produces no change and recursion is invisible -- which
        // is what the published prior art still does. rsp at the dispatch falls
        // exactly one step per VBA call and is stable within a frame (8 descents
        // for T_Rec(8), 100 for T_Deep(100)), so (trailer, sp) identifies an
        // ACTIVATION and recursion becomes an ordinary push.

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

        // DOES THIS EXIT END THE PROCEDURE? Not all do -- a GoSub's `Return`
        // fires one mid-activation, and closing on it cuts the activation in
        // half. vbaboundary.h owns that distinction.
        //
        // Both non-Yes answers close nothing and both are counted, because "this
        // exit was a Return" and "the opcode could not be read" are different
        // facts. Not closing is the safe direction: a missed close is corrected
        // by the next prologue, the stack pointer or the flush, while a wrong
        // close writes a row asserting a call VBA never made.
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
        // `End` TEARS THE WHOLE VBA SESSION DOWN, and it is the only moment those
        // activations are observably over.
        //
        // WHY THIS HOOK EXISTS AT ALL. `End` fires no exit opcode for any frame it
        // kills, so the frames could only be closed by the stack-pointer backstop --
        // which needs a later statement at a HIGHER rsp. A rebuild evaluates the
        // sheet twice, and when the second pass starts DEEPER no such statement ever
        // arrives: the new activations nested under the dead ones and the trace
        // reported depth 8 where 4 was right. rsp alone cannot tell "deeper because
        // nested" from "deeper because the last chain was abandoned"; the End opcode
        // can, because it means exactly one thing.
        //
        // THIS THREAD ONLY, because the shadow stack is TLS. Another thread's frames
        // are closed by its own backstop or the flush -- closing them from
        // here would be writing rows for a stack we are not synchronised with.
        //
        // TICKS ARE A MEASUREMENT HERE, unlike the other two abrupt closes: the
        // opcode fires when the session dies, not at whatever happened next.
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

    void FlushOpenFrames()
    {
        // t_state is the storage itself, not a pointer to it, so a thread that
        // never traced has generation 0 and depth 0 and this is a no-op for it.
        ThreadState* s = &t_state;
        const std::uint64_t now = Now();
        // A frame this thread is still inside sits above the current stack with its
        // trailer in place. A leftover from an earlier unwind has lost that stack.
        const std::uint64_t here = reinterpret_cast<std::uint64_t>(_AddressOfReturnAddress());
        while (s->depth > 0)
        {
            const Frame& top = s->stack[s->depth - 1];
            // Still running only if it sits above our own stack AND its trailer is in
            // place; a leftover from an earlier unwind fails the second test.
            const bool running = top.sp > here && FrameStillLive(top);
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
