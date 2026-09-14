// WHERE DOES A VBA ACTIVATION BEGIN AND END? Two questions, and nothing else:
//
//     is this statement the start of an activation?
//     does this exit opcode end one?
//
// vbatrace.cpp asks them and manages the shadow stack without knowing how they
// are answered. THAT SEAM IS THE POINT: when a better mechanism turns up this is
// the file replaced, and nothing above it changes.
//
// START is the statement whose instruction IS the first byte of the procedure --
// a prologue, not statement 1, which is why a loop never returns to it (459
// activations, 28 shapes, three entry routes, against a counter VBA increments
// itself: exactly one per activation. `Do While`, `While...Wend`, `For` and a
// top-of-body label all branch back to offset 6 or later, never to 0).
// The obvious alternative, opening on the STACK POINTER, fails because GoSub
// moves rsp within one activation.
//
// A THIRD WAY AN ACTIVATION ENDS, and it is not this file's: `End` kills the
// whole session and fires no exit opcode at all, so nothing here can see it.
// It is answered by hooking the `End` opcode itself (slot 619), which
// closes the chain directly rather than being read off the interpreter here.
//
// END is an exit opcode, EXCEPT the GoSub `Return` (620) and the two heap-type
// cleanup exits (1662, 1663, which fire before the real end at 1664 for
// String/Object/Variant returns). Closing on ANY exit fails because the family
// is statement EPILOGUES and a UDF that raises never reaches one; position
// cannot substitute (a `Helper:` block sits at the END of a procedure) and
// vbaderive.cpp's repeat-groups group by shared HANDLER, not meaning.
//
// HOW STRONG EACH CLAIM IS -- they are not equal:
//   corpus-grade : a slot index means the same opcode on every build, and 620
//                  is an exit slot at all. 41 VBE7 binaries, 2012-2026.
//   this build   : 620 is the GoSub `Return` and 1662/1663 are cleanup. One
//                  build, 7.1.11.58; if it changes, activation counts double.
//   this build   : the prologue is one-per-activation. The p-code is compiled
//                  INTO the workbook, so an old workbook is the untested case.
//
// NOT ESTABLISHED: only a handful of the exit slots have been seen firing, so
// `ProcedureEnd` closes on everything except the known mid-procedure ones -- a
// DENY-list, because an allow-list would silently stop closing for a procedure
// kind nobody tested.
//
// WHAT WOULD REPLACE THIS: EbGetCallstackCount / EbGetCallstackFunction, VBE7's
// own answer, resolvable from Microsoft's public PDBs but not exported -- or
// Excel's COM entry into VBA, which has .pdata by necessity. Either would collapse this file to one call.
//
// NO STATE, NO COUNTERS, NO SIDE EFFECTS: both functions are pure reads of the
// interpreter's registers and bytecode. The caller counts the outcomes, because
// a policy that keeps its own books is one you cannot swap out.
#pragma once
#include <cstdint>

namespace vba
{
    // `Unavailable` is a THIRD answer, never folded into `No`: RSI or ProcSize
    // could not be read, so the question was not answered rather than answered
    // negatively. The caller degrades and says so.
    enum class Activation { No, Yes, Unavailable };
    Activation ActivationStart(std::uint64_t trailer, std::uint64_t savedRegs);

    // Does this exit dispatch end the procedure?
    //   Yes         -- close the frame here.
    //   No          -- a GoSub `Return`; the activation continues.
    //   Unreadable  -- the opcode could not be read. Treated as No by the
    //                  caller, deliberately: a missed close is corrected by the
    //                  next prologue, the stack pointer, or the flush, whereas
    //                  a wrong close writes a row asserting a call VBA never
    //                  made. Counted apart from No so the two never merge.
    enum class Ending { No, Yes, Unreadable };
    // `opOut` receives the exit opcode. The exits are TYPED -- a procedure
    // leaves through a slot chosen by its declared return kind (vbaretdecode.h) --
    // so the number names the return type at the one moment the result can be
    // read.
    Ending ProcedureEnd(std::uint64_t savedRegs, std::uint16_t* opOut = nullptr);
}
