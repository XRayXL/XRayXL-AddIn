// Reads a VBA procedure's own bytecode to recover its argument types. The frame gives values
// and says nothing about what they are; the p-code is typed, so the type comes from the opcode.
//
// The trailer trails the code: trailer+0x0C is ProcSize, so the bytecode is [trailer-ProcSize,
// trailer). Walking it needs instruction lengths, which are a pinned constant of the opcode set
// (kSigLength).
//
// Fail-closed: the walk stops at an exit opcode or exactly at ProcSize, and resynchronises on
// the next beginning-of-statement after an unknown length. A parameter the body never reads
// emits no load, so it has no recoverable type.
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
    //
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

        // Lengths no real compiler output has exercised: pinned from the handler alone, and
        // absent from the offline corpus. Counting a step through one says which to check next.
        // One array index on the hot path.
        bool          unverified[kMax] = {};
        std::uint32_t unverifiedCount  = 0;

        // Slots whose length DEPENDS ON THE OPERAND: `len` is the base and the
        // count word at +2 adds `varUnit` bytes per unit. 0 = fixed length.
        std::uint8_t  varUnit[kMax] = {};
        // Which slots can name a parameter at all. A parameter lives at `[R14 + 8n]`, so an
        // opcode whose handler never computes an address from R14 cannot load or store one.
        // Attributing on operand shape alone would read the next instruction's bytes for a
        // two-byte opcode. Derived once at arm.
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

        // The opcode the walk stopped on, and with it the declared return type
        // (ExitReturnKind). Recorded because the argument walk needs it: a Variant-returning
        // Function is handed the caller's result VARIANT in an argument slot. 0 when the walk
        // did not reach an exit.
        std::uint16_t exitOp     = 0;

        // THE EXIT INSTRUCTION'S OPERAND. For a class or form Function this is
        // the byte offset of the trailing slot holding the result, so the claim
        // "the last slot is not a parameter" can be CHECKED against the
        // bytecode rather than inferred from the opcode alone.
        std::int32_t  exitOperand = 0;

        // Did the walk not read the whole body: stopped early, or resynchronised past a
        // statement? Either way a `?` may be a skipped load rather than an unread parameter.
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

    // The same ranking as a warning, or "" when no walk was stopped. The walk cannot step over
    // an opcode with no known length, so it resynchronises and skips everything between, typed
    // loads included.
    const char*   PcodeStopWarning();

    // Opcodes whose length was USED and after which the walk broke: a length
    // that is WRONG rather than missing, the worse defect. Carries the
    // clean-walk fraction, the table's health number.
    const char*   PcodeSuspectWarning();
    // Walks whose last statement did not end at ProcSize, by exit; "" when there were none.
    const char*   PcodeClosureWarning();
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

    // Where a Function stores its result, for the return decoder. The result is at [R14-8], a
    // Variant's at [R14-0x18], and the instruction that wrote it carries the type.
    //
    // Scanned, not walked, so the return path stays independent of the length table. Every
    // distinct (opcode, operand) pair in order of first occurrence, named or not; `max` bounds
    // the output.
    int ReturnStoreCandidates(std::uint64_t trailer, std::uint16_t* ops,
                              std::int32_t* offs, int max);

    // The declared type a STORE opcode carries, via store = load + 32. Returns
    // nullptr rather than a guess.
    const char* PcodeStoreTypeName(std::uint32_t storeOp);

    // Walks the procedure behind `trailer` and fills in what it declares. `maxArg` is the true
    // number of argument slots, from argSz; a recovered index above it is discarded. Pass 0
    // when unknown.
    bool ReadArgTypes(std::uint64_t trailer, ArgTypes& out, int maxArg = 0);
}
