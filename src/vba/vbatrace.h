// The VBA tracer: identity and the shadow stack.
//
// The thunk calls in with the interpreter's rsp, and the p-code trailer -- the
// RTMI identifying the running procedure -- is at rsp+0xB8. That one pointer
// yields the call tree, nesting depth, recursion and per-procedure timings. A
// trailer is per-procedure, so it is the accounting key; the NAME is walked
// from it on every frame push (vbaidentity.h).
//
// THE HOT PATH, per statement: a few guarded reads, a compare, and usually a
// return. No allocation, no lock, no call into VBE7. State is TLS, and the
// procedure table is fixed-size and open-addressed, so a new procedure costs a
// bounded probe and never allocates.
//
// RECURSION IS WHY THE SHADOW STACK EXISTS: a recursive call has the SAME
// trailer as its caller, so nothing in the interpreter distinguishes the two.
// The published prior art tracks a single "current procedure", to which
// recursion is invisible.
#pragma once
#include <cstdint>
#include <string>

namespace vba
{
    // Called from vbathunk.asm.
    //   dispatchSp = the interpreter's rsp at the dispatch site
    //   savedRegs  = top of the thunk's saved register block; the interpreter's
    //                registers lie below it, r14 (the VBA frame base) at -0x68
    //                (vbaregs.h is the one authority for these offsets)
    extern "C" void XRayVbaOnStatement(std::uint64_t dispatchSp, std::uint64_t savedRegs);
    extern "C" void XRayVbaOnExit(std::uint64_t dispatchSp, std::uint64_t savedRegs);
    // The RAISE opcode. Fires four times per Err.Raise, so the recorder
    // dedupes. Patched only when the raise slot verified (vbaderive.h).
    extern "C" void XRayVbaOnRaise(std::uint64_t dispatchSp, std::uint64_t savedRegs);

    struct Totals
    {
        std::uint64_t statements = 0;   // beginning-of-statement opcodes seen
        std::uint64_t exits = 0;        // exit opcodes seen
        std::uint64_t transitions = 0;  // times the running procedure changed
        std::uint64_t procedures = 0;   // distinct trailers seen
        std::uint64_t recursions = 0;   // pushes onto a stack already holding that trailer
        // Capped by the shadow stack; `deepestSeen` is the depth VBA reached.
        std::uint32_t maxDepth = 0;

        // A LOWER BOUND on how deep the VBA stack actually got, counting
        // activations the shadow stack had no room for; equals `maxDepth` when
        // nothing overflowed. A bound and not a number because beyond the cap
        // no rsp is stored to match a return against, so one unwind abandoning
        // several capped activations decrements by one and the peak is
        // undercounted. Never OVER-counted, so it is reported as ">=".
        std::uint32_t deepestSeen = 0;
        std::uint64_t faults = 0;       // guarded reads that faulted

        // Caught by the hook's own SEH frame. Distinct from `faults`, a
        // guarded read declining: this is the tracer itself going wrong inside
        // a VBA thread, and it is what trips the circuit breaker.
        std::uint64_t hookFaults = 0;
        std::uint64_t hookReentries = 0;   // hooks skipped: the thread was already inside one
        bool          tripped = false;  // breaker open: the tracer has stood down

        // Apart from other faults, because catching one means the thread's
        // guard page is gone and tracing on it risks an unrecoverable fault
        // later, in unrelated code.
        std::uint64_t stackOverflows = 0;

        // A failed read of the thunk's saved-register block: the ABI shared with
        // hand-written assembly has broken. Counted apart from "no frame base".
        std::uint64_t regReadFailures = 0;
        std::uint64_t overflows = 0;    // pushes past the shadow stack's capacity
        std::uint64_t tableFull = 0;    // procedures we could not record
        std::uint64_t unmatchedExits = 0; // exit with nothing on our stack

        // `exitClosed` is the PRIMARY way a frame ends. `exitNoMatch` means the
        // exit did not match the frame on top -- normal when the stack pointer
        // already closed it (an error unwind), suspicious otherwise.
        std::uint64_t exitClosed = 0;
        std::uint64_t exitNoMatch = 0;

        // framesOpened == framesClosed after a flush is THE invariant for error
        // handling and teardown: a frame we failed to follow shows up as never
        // closed. An unhandled error and `End` both abandon frames without
        // firing an exit opcode, so neither is caught by the exit count.
        std::uint64_t framesOpened = 0; // frames we pushed
        std::uint64_t framesClosed = 0; // frames we popped

        // How rsp moved while the trailer stayed the same. It falls exactly one
        // step per VBA call and is stable within a frame (8 descents for
        // T_Rec(8), 100 for T_Deep(100)), which is what makes (trailer, sp) an
        // ACTIVATION. The cheapest way to see the fast path working: spSame
        // should dominate.
        std::uint64_t sameTrailerSameSp = 0;
        std::uint64_t sameTrailerDeeper = 0;   // rsp fell: a deeper frame
        std::uint64_t sameTrailerShallower = 0;// rsp rose: returned

        // THE ACTIVATION BOUNDARY, from the p-code instruction pointer. A
        // procedure's bytecode is [trailer - ProcSize, trailer) and RSI-2 is the
        // current instruction, so an instruction at the FIRST byte is the
        // prologue -- which runs exactly once per activation (459 activations,
        // 28 shapes, three entry routes; no loop branches back to offset 0).
        //
        // It is what stops a UDF that RAISES from swallowing the next call: it
        // never reaches its epilogue, so the next prologue closes its frame.
        std::uint64_t ipEntries = 0;      // prologue statements seen
        std::uint64_t ipStaleClosed = 0;  // frames a prologue found still open
                                          // at the same rsp -- the defect, counted

        // The rule could not be evaluated: RSI unreadable, ProcSize absent or
        // insane. NOT merged with "no entry seen" -- while this is non-zero the
        // boundary has degraded to the rsp/exit-opcode heuristic.
        std::uint64_t ipUnavailable = 0;

        // ONLY A PROLOGUE OPENS A FRAME; the stack pointer still closes them,
        // which it does correctly. This counts an rsp change inside a procedure
        // already on the stack -- GoSub, Return, a handler moving the stack.
        std::uint64_t ipIntraProc = 0;

        // A frame opened WITHOUT a prologue: a procedure already mid-flight
        // when tracing armed. It is opened anyway -- attributing its statements
        // to nothing would be worse -- but its start time and parent are
        // guesses, and this is what says so. Zero in a clean session.
        std::uint64_t ipLateOpen = 0;

        // NOT EVERY EXIT OPCODE ENDS A PROCEDURE: the "exit" family is
        // statement epilogues and a GoSub `Return`
        // fires one mid-activation. `exitOpUnreadable` closes nothing rather
        // than guessing -- a missed close is corrected by the next prologue, the
        // stack pointer or the flush, while a wrong one writes a row asserting a
        // call VBA never made.
        std::uint64_t exitNotProcedureEnd = 0;
        std::uint64_t exitOpUnreadable = 0;

        // WHO CALLED THIS VBA PROCEDURE. Once per ACTIVATION, never per
        // statement: xlfCaller is a call INTO Excel from a traced thread, which
        // is affordable at the first granularity and not at the second.
        //
        // Three counters for three different facts. `callerCell` is a real cell
        // on a real sheet. `callerOther` is a caller that is not a cell and
        // never could be -- a button names the OBJECT, the macro dialog gives
        // #REF! -- which are answers, not failures. `callerUnavailable` is Excel
        // declining, the only one saying something is wrong with where we asked
        // from.
        std::uint64_t callerCell = 0;
        std::uint64_t callerOther = 0;
        std::uint64_t callerUnavailable = 0;
        std::uint64_t callerFaults = 0;   // the guarded call itself faulted

        // RETURN VALUES and the declines beside them (vbaretdecode.h): counting
        // refusals keeps "we do not read returns" distinguishable from "we read
        // one and it was empty".
        std::uint64_t returnsRead = 0;
        std::uint64_t returnsDeclined = 0;
        // An exit opcode ExitReturnKind has no mapping for -- NOT a declined
        // read, since the decoder was never called. The action log names the
        // opcodes, because the value is what has to be mapped.
        std::uint64_t returnsUnmapped = 0;

        // ---- ByRef ARGUMENTS A PROCEDURE CHANGED -------------------------
        //
        // ByRef is VBA's DEFAULT, so a procedure filling in its caller's
        // variable is ordinary code and an entry-only args column tells half the
        // story. The exit re-reads them and reports them WHEN THEY MOVED.
        //
        // Four counters because collapsing them would make "nothing changed"
        // indistinguishable from "we never looked".
        //   Eligible : the signature had a ByRef parameter. A superset.
        //   Changed  : re-read, and different -- the rows carrying exit args.
        //   Same     : re-read, and identical. A fact about the procedure.
        //   Declined : eligible, but the re-read produced nothing. NOT "same".
        std::uint64_t byrefEligible = 0;
        std::uint64_t byrefChanged  = 0;
        std::uint64_t byrefSame     = 0;
        std::uint64_t byrefDeclined = 0;

        // ---- ERRORS: where one was thrown, and who caught it -------------
        //
        // `raises` counts firings AFTER dedupe -- the opcode fires four times
        // per Err.Raise, so the raw count is meaningless. `threw`, `unwound` and
        // `handled` are activations: threw == handled where every error was
        // caught, and a shortfall is an error that reached the top.
        std::uint64_t raises = 0;
        std::uint64_t raisesDeduped = 0;   // the extra firings, discarded
        std::uint64_t threw = 0;
        std::uint64_t unwound = 0;
        std::uint64_t handled = 0;
        // A raise VBE7 resolved IN PLACE. Ordinary object-model VBA -- a cell
        // write, For Each, Application.OnTime -- reaches the same raise opcode
        // (497) as Err.Raise, but its frame keeps running where a real unhandled
        // throw unwinds. Counted so it is not read as the throw it is not.
        std::uint64_t raisesBenign = 0;
        // A raise whose frame was not open yet, held until it was. Ordinary for
        // a UDF beginning with Err.Raise, not a fault.
        std::uint64_t errLateOpen = 0;
        // A raise with an empty shadow stack: nothing to attribute it to, so
        // counted rather than guessed at.
        std::uint64_t errNoFrame = 0;
        // An error that unwound past every frame and left VBA -- into a cell as
        // #VALUE!, or into whatever called the macro. `threw` exceeds `handled`
        // by exactly this.
        std::uint64_t errEscaped = 0;

        // ---- WAS THE DURATION MEASURED, OR BOUNDED? ----------------------
        //
        // Only the exit opcode fires when an activation actually ends. The
        // stack-pointer backstop fires at the NEXT statement and the flush at
        // disarm, so those rows carry an UPPER BOUND. Counted so a session can
        // say how much of its timing is bounded without reading every row.
        std::uint64_t closedByBackstop = 0;
        std::uint64_t closedByFlush = 0;
        // `End` is the one abrupt close whose ticks ARE a measurement: the
        // opcode fires at the moment the session dies, so the frame is closed
        // when it actually ended rather than whenever the next statement
        // happened to arrive. Counted apart for exactly that reason.
        std::uint64_t closedByEnd = 0;
    };

    // Wait until no thread is inside a hook. False if it did not drain in
    // time, in which case the CALLER MUST NOT reset anything: stale state is
    // recoverable, resetting it under a live thread is not.
    bool   WaitForHooksQuiet(unsigned milliseconds);

    // Only depth-1 frames emit rows. Latched at arm from DEPTH(VBA); the
    // totals count everything regardless.
    void   SetEmitTopLevelOnly(bool topOnly);

    // Whether each frame asks Excel WHO called it -- an Excel12 round-trip on
    // the traced thread, inside the span being timed. FALSE buys it back and
    // the row says `not-asked`.
    // Latched at arm, so the hot path never sees a setting change under it.
    void   SetCapture(bool args, bool retval);

    // ---- reading the procedure table WHILE ARMED ------------------------
    // Report() reads it at disarm, after the hooks have drained;
    // XRayXL_GetTraceSummary reads it LIVE from a worksheet, so it copies out
    // under the publish flag. Entries are only ADDED and counters only rise, so
    // a racing reader sees a stale count, never garbage -- and a name being
    // republished is reported as its trailer address, not a half-written string.
    int    ProcTableSize();
    bool   ProcSnapshot(int slot, unsigned long long& trailer,
                        unsigned long long& calls,
                        char* module, int modCap, char* function, int fnCap);

    void   ResetTracing();

    // Close this thread's open frames, so the trace does not end with entry
    // rows that never got an exit. Only the calling thread's stack is
    // reachable; other threads' frames stay open and visible as unmatched
    // entries rather than being silently invented.
    void   FlushOpenFrames();

    Totals ReadTotals();

    // One line per procedure, most time first. Off the hot path.
    std::string Report(int maxRows = 25);

    // Exit opcodes that ended an activation, that ExitReturnKind cannot name
    // and the store-scan fallback did not rescue. The return value was left
    // EMPTY, which is correct and also invisible; this makes it visible.
    std::string ReturnTypeUnknownWarning();
    std::string TotalsLine();
}
