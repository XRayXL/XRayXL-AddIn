// Reading a VBA procedure's own bytecode, to recover its ARGUMENT TYPES.
//
// The frame gives argument VALUES and says nothing about what they are: eight
// bytes reading 0x40A5790000000000 are a Double, a large Long, or a pointer.
// The p-code is TYPED, so the type comes from the OPCODE, not the bytes.
//
// The trailer *trails* the code: trailer+0x0C is `ProcSize`, so the bytecode is
// [trailer-ProcSize, trailer).
//
// WALKING IT NEEDS INSTRUCTION LENGTHS -- a 2-byte opcode plus a variable
// operand. Those lengths are a PINNED CONSTANT of the opcode set measured across
// 42 VBE7 builds (kSigLength); nothing is derived at arm time.
//
// FAIL-CLOSED: the walk stops at an exit opcode or exactly at ProcSize, and
// resynchronises on the next beginning-of-statement after an unknown length or
// a desync -- never stepping over what it cannot size. Types found while
// synchronised are trustworthy; types after a stop are absent.
//
// A parameter the body never reads emits no load, so it has no recoverable type
// at all.
#pragma once
#include <cstdint>
#include <string>
#include "vbaderive.h"

namespace vba
{
    // Two pairs share an opcode, correctly: Integer and Boolean are both
    // two-byte I2, and a VBA Date *is* a Double. Both follow from the
    // OLE Automation types, which is a reason to believe the table.
    const char* PcodeTypeName(std::uint32_t loadOpcode);

    // An opcode that reaches a parameter slot and genuinely carries NO type --
    // 671 is an address-of, not a load. `?opNNN` would tell a reader to add a
    // table row that cannot help; a bare `?` is the truth.
    // [measured: the handler]
    bool        PcodeCarriesNoType(std::uint32_t op);

    // Instruction lengths, one per dispatch slot, from the pinned table.
    struct PcodeLengths
    {
        static constexpr int kMax = 2048;   // >= the 1700 the table actually has
        std::uint8_t  len[kMax]  = {};      // 0 = not derived
        std::uint32_t slots      = 0;
        // What the table put in place. Nothing to cross-check against: that
        // question is asked offline, against a
        // corpus of builds rather than one machine.
        std::uint32_t pinned     = 0;      // entries written
        // Slots pointing at the shared invalid-opcode handler. Landing on one
        // means the walk is ALREADY LOST -- these are not instructions -- so
        // whatever length got there is wrong. Distinct from `len[i] == 0`, which
        // says only that a real instruction's length is unknown.
        bool          invalid[kMax] = {};
        std::uint32_t invalidCount  = 0;

        // LENGTHS NO RUNNING CODE HAS EVER EXERCISED. A slot the offline
        // corpus has walked past had its length confirmed by real compiler
        // output; one that is pinned and absent from it rests on the handler
        // alone, the weaker evidence. Counting a step through one is how a real workbook --
        // which runs shapes no generator writes -- tells us which to measure
        // next. Class T: one array index on the hot path.
        bool          unverified[kMax] = {};
        std::uint32_t unverifiedCount  = 0;

        // Slots whose length DEPENDS ON THE OPERAND: `len` is the base and the
        // count word at +2 adds `varUnit` bytes per unit. 0 = fixed length.
        std::uint8_t  varUnit[kMax] = {};
        // WHICH SLOTS CAN NAME A PARAMETER AT ALL. A parameter lives at
        // `[R14 + 8n]` and nowhere else, so an opcode whose handler never
        // computes an address from R14 cannot load or store one. Of 1700 slots,
        // 677 are the shared invalid handler, and of the 1023 real ones only 356
        // address the frame.
        //
        // Attributing on OPERAND SHAPE alone -- positive, a multiple of eight --
        // reads the NEXT instruction's bytes for the 503 non-frame opcodes that
        // are two bytes long.
        //
        // Derived once at arm: class T, so the hot path costs an array index.
        bool          framesR14[kMax] = {};
        std::uint32_t framedCount     = 0;
        std::uint32_t scanMicros      = 0;   // what that pass cost, once per arm
        bool          ok         = false;
    };

    // Fill the table from the pinned constants, bounded by the verified
    // SlotSet's size. Nothing is derived here; derivation is offline.
    bool PinPcodeLengths(const Image& img, const SlotSet& slots,
                         PcodeLengths& out);

    // The lengths for the module armed in this process, or nullptr.
    const PcodeLengths* ArmedLengths();
    void SetArmedLengths(const PcodeLengths& l);
    void ClearArmedLengths();
    // Latched at every arm from XRAYXL_DIAG, whether or not lengths get pinned.
    void SetPcodeDiagnostics(bool on);

    // Argument types for one procedure, indexed by frame SLOT: argument n is
    // at [R14 + 8n], and a ByVal Variant occupies three slots. A null entry
    // means the type was not recoverable -- usually because the body never
    // reads that parameter.
    struct ArgTypes
    {
        // Indexed by SLOT, so it tracks vbaargs' kMaxSlots: 60 parameters can
        // occupy 181 slots when they are `ByVal Variant`.
        static constexpr int kMax = 192;
        const char*   name[kMax] = {};   // nullptr where unknown
        // The opcode seen at the slot, named or not: when a type is missing
        // this is the ONLY thing saying which opcode to add, so a gap reports
        // itself rather than rendering an anonymous "?".
        std::uint16_t op[kMax]   = {};

        // THE OPCODE THE WALK STOPPED ON, and with it the declared return type
        // -- the compiler picks the exit from what the procedure returns (952
        // for a Variant, 635 for a Sub, 631 for an object reference), which
        // `ExitReturnKind` already maps. It is recorded here because the
        // ARGUMENT walk needs it: a Variant-returning Function is handed the
        // caller's result VARIANT in an argument slot, and at entry there is no
        // exit row to ask. 0 when the walk did not reach an exit.
        std::uint16_t exitOp     = 0;

        // THE EXIT INSTRUCTION'S OPERAND. For a class or form Function this is
        // the byte offset of the trailing slot holding the result, so the claim
        // "the last slot is not a parameter" can be CHECKED against the
        // bytecode rather than inferred from the opcode alone.
        std::int32_t  exitOperand = 0;

        // Did the walk NOT read the whole body -- stopped early, or
        // resynchronised past a statement and still reached the exit? Either
        // way a `?` may be a skipped load rather than an unread parameter.
        // (Carrying only how the walk ended once let the second shape report
        // itself complete.)
        //
        // The return type comes from ReturnStoreCandidates(), not from here.
        bool          partial    = false;
    };

    enum class PcDecline
    {
        NoLengths, NoProcSize, ProcSizeInsane, CodeUnreadable,
        Desynced, UnknownOpcode, Count_
    };
    const char*   PcDeclineName(PcDecline d);
    void          ResetPcodeCounts();
    const char*   PcodeLine();

    // The (opcode, operand) stream of ONE procedure whose parameters did not
    // all resolve -- the instrument saying whether an unrecovered type is a gap
    // in the TABLES or in the WALK, which a `?` signature cannot. Empty when
    // everything resolved.
    const char*   PcodeUntypedDump();

    // The opcodes that actually stopped walks, most frequent first -- the
    // ranking that decides whether more length work is worth doing, since an
    // opcode that never occurs costs nothing however unlengthed.
    const char*   PcodeStopLine();

    // THE SAME RANKING, AS A WARNING, or "" when no walk was stopped. An
    // opcode with no known length is the strongest sense of "we did not
    // understand this instruction": the walk cannot step over it, so it
    // resynchronises and SKIPS EVERYTHING BETWEEN, typed loads included -- which
    // is how a parameter reads `(?)` and is reported as one the body never uses.
    const char*   PcodeStopWarning();

    // Opcodes whose length was USED and after which the walk broke: a length
    // that is WRONG rather than missing, the worse defect. Carries the
    // clean-walk fraction, the table's health number.
    const char*   PcodeSuspectWarning();
    // LENGTHS THIS SESSION USED THAT NOTHING HAS EVER CONFIRMED, ranked.
    // Empty when the session met none, which is the ordinary case: the point
    // of the line is that a real workbook found one and we did not.
    const char*   PcodeUnverifiedWarning();

    // The table's health as two numbers: procedures walked, and those walked
    // cleanly from offset 0 to an exit with no resynchronisation. A correct
    // table walks every procedure cleanly, so anything less is a defect and
    // this is how a caller outside the log can see it.
    void          PcodeHealth(long long& walks, long long& clean);

    // What the R14 frame gate refused; empty when it refused nothing.
    const char*   PcodeNotFramedWarning();

    // Write every distinct procedure walked this session as raw bytes, for the
    // offline length solver. XRAYXL_DIAG only; call at disarm, never in a hook.
    std::string   WritePcodeCorpus(const std::wstring& path);

    // Where a Function stores its result, for the return decoder; retdecode
    // owns what the store opcodes mean. The result is at [R14-8], a Variant's
    // at [R14-0x18], and the frame gives bytes, not a type: the instruction
    // that wrote it carries the type, as a load does for an argument.
    //
    // Found by scanning, not walking: the walk resynchronises past the store
    // in most procedures, and the return path stays independent of the length
    // table so a wrong length costs a parameter type, never a `ret`. The scan
    // requires the 4-byte result offset with an opcode word in front of it.
    //
    // Every distinct (opcode, operand) pair, in order of first occurrence,
    // named or not; `max` bounds the output. The caller judges disagreement.
    int ReturnStoreCandidates(std::uint64_t trailer, std::uint16_t* ops,
                              std::int32_t* offs, int max);

    // The declared type a STORE opcode carries, via store = load + 32. Returns
    // nullptr rather than a guess.
    const char* PcodeStoreTypeName(std::uint32_t storeOp);

    // Walk the procedure behind `trailer` and fill in what it declares.
    // `maxArg` is the true number of argument slots, from argSz in the trailer;
    // a recovered index above it is PROVABLY wrong -- the frame is not that big
    // -- and is discarded, which turns a bad resync landing from a plausible
    // wrong type into nothing at all. Pass 0 when unknown.
    bool ReadArgTypes(std::uint64_t trailer, ArgTypes& out, int maxArg = 0);
}
