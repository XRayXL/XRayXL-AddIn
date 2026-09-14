// Patching the p-code dispatch table: statement, exit, raise and End slots swapped
// for thunks that count and jump straight on to the original handler.
//
// A TABLE PATCH IS SAFER THAN A CODE PATCH. Nothing in the instruction stream is
// rewritten, so Control Flow Guard, CET shadow stacks and unwind metadata are
// untouched, and disarming is a pointer write back rather than a byte restore.
// An 8-byte aligned pointer store is atomic on x64, so a thread dispatching
// concurrently sees either the old handler or the new one, and both work.
//
// THE ONE THING TO UNDERSTAND ABOUT THE THUNK: a handler is entered by
// `jmp qword ptr [rbx+rax*8]`, NOT by a call. There is no return address on the
// stack and the handler owns the interpreter's live registers, so the thunk must
// leave the stack byte-identical, preserve every register and flag, and
// tail-jump. It touches nothing but its own page.
//
// OFF BY DEFAULT: VBE7 is patched only when the VBA source is set to DEPTH=TOP
// or DEPTH=ALL, so installing the add-in patches nothing unless the
// user asks -- and the baseline half of a cost measurement is an Arm with the
// mode OFF.
#pragma once
#include <cstdint>
#include <string>
#include <vector>

namespace vba
{
    // The mode setters refuse while it is patched, because the modes are
    // read once, at arm.
    bool IsVbaArmed();

    // Derive, verify, and patch. Refuses unless the derivation verified.
    // Returns a line for the action log.
    std::string ArmCounting();

    // Restore every original pointer; safe when never armed. The thunk page is
    // intentionally NOT freed -- see the .cpp.
    std::string DisarmCounting();

    // WHAT THE TRACER DID NOT UNDERSTAND THIS SESSION, at WARNING level: an
    // opcode with no known LENGTH (the walk skipped a statement), one reaching a
    // parameter slot with no TYPE (`?opNNN`), and an exit opcode with no RETURN
    // mapping (an empty `ret`). Each is a hole a reader would otherwise mistake
    // for an absence of evidence -- "the body never uses this parameter", "this
    // Sub has no result". Empty is the state a fuzz run asserts.
    std::vector<std::string> UnknownOpcodeWarnings();

    // The p-code diagnostics for the disarm report, at their proper log
    // levels: the warnings above, the clean-walk health check, unverified
    // lengths, the DIAG corpus file, and the frame gate's refusals.
    void LogDisarmDiagnostics();
}
