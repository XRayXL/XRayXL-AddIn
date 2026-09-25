// Patches the p-code dispatch table, swapping slots for thunks that count and jump on.
// No instruction is rewritten, so CFG, CET and unwind metadata are untouched; an aligned
// 8-byte store is atomic on x64, so a concurrent dispatch sees either handler.
// Handlers are entered by jmp, not call, so the thunk must preserve the stack, every register
// and flag. Patched only when the VBA source is DEPTH=TOP or DEPTH=ALL.
#pragma once
#include <cstdint>
#include <string>
#include <vector>

namespace vba
{
    // The mode setters refuse while patched, because the modes are read once, at arm.
    bool IsVbaArmed();

    // Refuses unless the derivation verified. Returns a line for the action log.
    std::string ArmCounting();

    // Safe when never armed. The stub page is never freed: a thread may still be inside a stub.
    // Returns the counts line; `procedures` gets the per-procedure line.
    std::string DisarmCounting(std::string& procedures);

    // After a contained fault: lets go of the arm/disarm gate if this thread held it, so the
    // next disarm can run. True if it was held.
    bool ReleaseArmGateHeldByThisThread();

    // XRAYXL_DIAG instrument: faults while holding the gate (XRayXL_FaultProbe "ARMGATE").
    void FaultWhileArmGateHeldForProbe();

    // What the tracer did not understand this session. Empty is the state a fuzz run asserts.
    std::vector<std::string> UnknownOpcodeWarnings();

    void LogDisarmDiagnostics();
}
