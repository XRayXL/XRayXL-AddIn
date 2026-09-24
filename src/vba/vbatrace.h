// The VBA tracer: identity and the shadow stack.
// The p-code trailer at rsp+0xB8 is the accounting key; the name is walked from it on every push.
// Per statement: a few guarded reads, a compare, usually a return. No allocation, lock or VBE7 call.
// The shadow stack exists for recursion: a recursive call has the same trailer as its caller.
#pragma once
#include <cstdint>
#include <string>

namespace vba
{
    // Called from vbathunk.asm. dispatchSp is the interpreter's rsp at dispatch; savedRegs is
    // the top of the thunk's saved register block (offsets in vbaregs.h).
    extern "C" void XRayVbaOnStatement(std::uint64_t dispatchSp, std::uint64_t savedRegs);
    // The same statement with a breakpoint on it, about to stop in the editor.
    extern "C" void XRayVbaOnBreakpoint(std::uint64_t dispatchSp, std::uint64_t savedRegs);
    // A `Stop` statement, the same pause written into the source. Its BoS has already run.
    extern "C" void XRayVbaOnStop(std::uint64_t dispatchSp, std::uint64_t savedRegs);
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

        // A lower bound, counting activations past the cap: one unwind abandoning several capped
        // activations decrements by one, so it is reported as ">=".
        std::uint32_t deepestSeen = 0;
        std::uint64_t faults = 0;       // guarded reads that faulted

        // The tracer itself faulting inside a VBA thread, unlike `faults`; trips the breaker.
        std::uint64_t hookFaults = 0;
        std::uint64_t hookReentries = 0;   // hooks skipped: the thread was already inside one
        bool          tripped = false;  // breaker open: the tracer has stood down

        // Apart, because the thread's guard page is gone and it risks a later unrecoverable fault.
        std::uint64_t stackOverflows = 0;

        // The ABI shared with the assembly has broken; counted apart from "no frame base".
        std::uint64_t regReadFailures = 0;
        std::uint64_t overflows = 0;    // pushes past the shadow stack's capacity
        std::uint64_t tableFull = 0;    // procedures we could not record
        std::uint64_t unmatchedExits = 0; // exit with nothing on our stack

        // `exitNoMatch` is normal when rsp already closed the frame (an error unwind),
        // suspicious otherwise.
        std::uint64_t exitClosed = 0;
        std::uint64_t exitNoMatch = 0;

        // Equal after a flush, or a frame was lost; an unhandled error and `End` fire no exit
        // opcode, so the exit count cannot show it.
        std::uint64_t framesOpened = 0;
        std::uint64_t framesClosed = 0;

        // How rsp moved while the trailer stayed the same; SameSp dominates when the fast path works.
        std::uint64_t sameTrailerSameSp = 0;
        std::uint64_t sameTrailerDeeper = 0;   // rsp fell: a deeper frame
        std::uint64_t sameTrailerShallower = 0;// rsp rose: returned

        // The prologue runs once per activation, which stops a UDF that raises from swallowing
        // the next call (vbaboundary.h).
        std::uint64_t ipEntries = 0;      // prologue statements seen
        std::uint64_t ipStaleClosed = 0;  // frames a prologue found still open at the same rsp

        // Apart from "no entry seen": while non-zero the boundary has degraded to rsp and exits.
        std::uint64_t ipUnavailable = 0;

        // An rsp change inside a procedure already on the stack: GoSub, Return, a handler.
        std::uint64_t ipIntraProc = 0;

        // A procedure already running at arm, opened without a prologue: its start time and
        // parent are guesses.
        std::uint64_t ipLateOpen = 0;

        // Each is a stretch of the call's duration spent in the debugger, not running.
        std::uint64_t breakpointStops = 0;
        // Counted apart from breakpointStops: a breakpoint is session state, a `Stop` is saved in
        // the workbook and ships to whoever opens it.
        std::uint64_t stopStatements = 0;

        // Not every exit opcode ends a procedure: a GoSub `Return` fires one mid-activation.
        // `exitOpUnreadable` closes nothing rather than guessing.
        std::uint64_t exitNotProcedureEnd = 0;
        std::uint64_t exitOpUnreadable = 0;

        // `callerOther` is a non-cell caller such as a button, an answer rather than a failure;
        // `callerUnavailable` is Excel declining.
        std::uint64_t callerCell = 0;
        std::uint64_t callerOther = 0;
        std::uint64_t callerUnavailable = 0;
        std::uint64_t callerFaults = 0;   // the guarded call itself faulted

        // Counting refusals keeps "not read" apart from "read, and empty".
        std::uint64_t returnsRead = 0;
        std::uint64_t returnsDeclined = 0;
        // RETVAL=FALSE, kept apart so "off" never reads as refusals.
        std::uint64_t returnsOff = 0;
        // No mapping for the exit opcode, so not a declined read: the decoder was never called.
        std::uint64_t returnsUnmapped = 0;

        // ByRef re-reads at exit, split so "nothing changed" is not "never looked":
        //   Eligible : the signature had a ByRef parameter
        //   Changed  : re-read and different; these rows carry exit args
        //   Same     : re-read and identical
        //   Declined : eligible, but the re-read produced nothing
        std::uint64_t byrefEligible = 0;
        std::uint64_t byrefChanged  = 0;
        std::uint64_t byrefSame     = 0;
        std::uint64_t byrefDeclined = 0;

        // ---- errors: where one was thrown, and who caught it ----
        std::uint64_t threw = 0;
        std::uint64_t unwound = 0;
        std::uint64_t handled = 0;
        // Left VBA into a cell or the macro's caller; `threw` exceeds `handled` by exactly this.
        std::uint64_t errEscaped = 0;
        // Chains Excel started while another VBA frame waited in DoEvents, reported at depth 1.
        std::uint64_t doEventsChains = 0;

        // The backstop fires at the next statement and the flush at disarm, so their ticks are
        // an upper bound; counted so a session can say how much timing is bounded.
        std::uint64_t closedByBackstop = 0;
        std::uint64_t closedByFlush = 0;
        // Apart, because `End` fires as the session dies, so its ticks are a measurement.
        std::uint64_t closedByEnd = 0;
    };

    // False if the hooks did not drain in time; the caller must then reset nothing, since
    // resetting under a live thread is not recoverable.
    bool   WaitForHooksQuiet(unsigned milliseconds);

    // Only depth-1 frames emit rows; the totals count everything regardless.
    void   SetEmitTopLevelOnly(bool topOnly);

    // Latched at arm, so the hot path never sees a setting change under it.
    void   SetCapture(bool args, bool retval);

    // Safe while armed: entries are only added and counters only rise, so a racing reader sees
    // a stale count; a name being republished comes back empty rather than torn.
    int    ProcTableSize();
    bool   ProcSnapshot(int slot, unsigned long long& trailer,
                        unsigned long long& calls,
                        char* module, int modCap, char* function, int fnCap);

    void   ResetTracing();

    // Only the calling thread's frames are reachable; other threads' stay visible as unmatched
    // entries rather than being invented.
    void   FlushOpenFrames();

    Totals ReadTotals();

    // One line per procedure, most time first. Off the hot path.
    std::string Report(int maxRows = 25);

    // Names exit opcodes whose return was left empty, which is correct but otherwise invisible.

    std::string ReturnTypeUnknownWarning();
    std::string TotalsLine();
}
