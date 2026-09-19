// The VBA tracer: identity and the shadow stack.
//
// The thunk calls in with the interpreter's rsp, and the p-code trailer identifying the running
// procedure is at rsp+0xB8. It is the accounting key; the name is walked from it on every frame
// push (vbaidentity.h).
//
// The hot path, per statement: a few guarded reads, a compare, and usually a return. No
// allocation, no lock, no call into VBE7.
//
// The shadow stack exists for recursion: a recursive call has the same trailer as its caller.
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

    // Called from the detour on VBE7's rtcDoEvents, on the thread entering and leaving it: a
    // procedure that opens while a frame waits in there was started by Excel, not called.
    void NoteDoEventsEnter();
    void NoteDoEventsLeave();

    struct Totals
    {
        std::uint64_t statements = 0;   // beginning-of-statement opcodes seen
        std::uint64_t exits = 0;        // exit opcodes seen
        std::uint64_t transitions = 0;  // times the running procedure changed
        std::uint64_t procedures = 0;   // distinct trailers seen
        std::uint64_t recursions = 0;   // pushes onto a stack already holding that trailer
        // Capped by the shadow stack; `deepestSeen` is the depth VBA reached.
        std::uint32_t maxDepth = 0;

        // A lower bound on how deep the VBA stack got, counting activations the shadow stack
        // had no room for. Beyond the cap no rsp is stored, so one unwind abandoning several
        // capped activations decrements by one. Never over-counted, so it is reported as ">=".
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

        // How rsp moved while the trailer stayed the same. It falls one step per VBA call and
        // is stable within a frame. spSame should dominate when the fast path is working.
        std::uint64_t sameTrailerSameSp = 0;
        std::uint64_t sameTrailerDeeper = 0;   // rsp fell: a deeper frame
        std::uint64_t sameTrailerShallower = 0;// rsp rose: returned

        // The activation boundary, from the p-code instruction pointer. RSI-2 is the current
        // instruction, so an instruction at the first byte of [trailer - ProcSize, trailer) is
        // the prologue, which runs once per activation. It stops a UDF that raises from
        // swallowing the next call.
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

        // Not every exit opcode ends a procedure: a GoSub `Return` fires one mid-activation.
        // `exitOpUnreadable` closes nothing rather than guessing.
        std::uint64_t exitNotProcedureEnd = 0;
        std::uint64_t exitOpUnreadable = 0;

        // Who called this VBA procedure, once per activation. `callerCell` is a real cell on a
        // real sheet. `callerOther` is a caller that is not a cell, such as a button or the
        // macro dialog: an answer, not a failure. `callerUnavailable` is Excel declining.
        std::uint64_t callerCell = 0;
        std::uint64_t callerOther = 0;
        std::uint64_t callerUnavailable = 0;
        std::uint64_t callerFaults = 0;   // the guarded call itself faulted

        // RETURN VALUES and the declines beside them (vbaretdecode.h): counting
        // refusals keeps "we do not read returns" distinguishable from "we read
        // one and it was empty".
        std::uint64_t returnsRead = 0;
        std::uint64_t returnsDeclined = 0;
        // Not asked for: RETVAL=FALSE. Kept apart, so "off" never reads as refusals.
        std::uint64_t returnsOff = 0;
        // An exit opcode ExitReturnKind has no mapping for -- NOT a declined
        // read, since the decoder was never called. The action log names the
        // opcodes, because the value is what has to be mapped.
        std::uint64_t returnsUnmapped = 0;

        // ByRef arguments a procedure changed. The exit re-reads them and reports them when
        // they moved. Four counters, so "nothing changed" is not "never looked":
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
        // Activations: threw == handled where every error was caught, and a
        // shortfall is an error that reached the top.
        std::uint64_t threw = 0;
        std::uint64_t unwound = 0;
        std::uint64_t handled = 0;
        // An error that unwound past every frame and left VBA -- into a cell as
        // #VALUE!, or into whatever called the macro. `threw` exceeds `handled`
        // by exactly this.
        std::uint64_t errEscaped = 0;
        // Chains Excel started while another VBA frame waited in DoEvents, reported at depth 1.
        std::uint64_t doEventsChains = 0;

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
