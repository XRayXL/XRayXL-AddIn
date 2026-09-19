// Patches the p-code dispatch table: statement, exit and End slots swapped for thunks
// that count and jump on to the original handler.
//
// Nothing in the instruction stream is rewritten, so CFG, CET and unwind metadata are
// untouched, and disarming is a pointer write back. An aligned 8-byte pointer store is atomic
// on x64, so a concurrent dispatch sees either handler.
//
// A handler is entered by `jmp qword ptr [rbx+rax*8]`, not a call, so the thunk must leave the
// stack byte-identical, preserve every register and flag, and tail-jump.
//
// VBE7 is patched only when the VBA source is DEPTH=TOP or DEPTH=ALL.
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

    // After a contained fault: lets go of the arm/disarm gate if this thread held it, so the
    // next disarm can run. True if it was held.
    bool ReleaseArmGateHeldByThisThread();

    // XRAYXL_DIAG instrument: faults while holding the gate (XRayXL_FaultProbe "ARMGATE").
    void FaultWhileArmGateHeldForProbe();

    // What the tracer did not understand this session, at WARNING level: an opcode with no
    // known length, one reaching a parameter slot with no type (`?opNNN`), and an exit opcode
    // with no return mapping. Empty is the state a fuzz run asserts.
    std::vector<std::string> UnknownOpcodeWarnings();

    // The p-code diagnostics for the disarm report, at their proper log
    // levels: the warnings above, the clean-walk health check, unverified
    // lengths, the DIAG corpus file, and the frame gate's refusals.
    void LogDisarmDiagnostics();
}
