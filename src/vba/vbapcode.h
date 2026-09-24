// Recovers argument types from a procedure's p-code: the frame holds untyped values, but each
// load opcode names its type. The bytecode is [trailer-ProcSize, trailer), ProcSize at +0x0C.
//
// Fail-closed: the walk stops at an exit or at ProcSize and resynchronises on the next
// statement after an unknown length. A parameter the body never reads has no recoverable type.
#pragma once
#include <cstdint>
#include <string>
#include "vbaderive.h"

namespace vba
{
    // Integer and Boolean share an opcode (both I2), as do Date and Double.
    const char* PcodeTypeName(std::uint32_t loadOpcode);

    // An opcode that reaches a parameter slot but carries no type (671 is an address-of), so
    // it renders a bare `?` rather than an `?opNNN` asking for a table row that cannot help.
    bool        PcodeCarriesNoType(std::uint32_t op);

    // Which address the callee receives: 751 the pointer the slot holds, 671 and 662 the slot's own.
    bool        PcodePassesHeldPointer(std::uint32_t op);
    bool        PcodePassesSlotAddress(std::uint32_t op);

    // Instruction lengths, one per dispatch slot, from the pinned table.
    struct PcodeLengths
    {
        static constexpr int kMax = 2048;   // >= the 1700 the table actually has
        std::uint8_t  len[kMax]  = {};      // 0 = not derived
        std::uint32_t slots      = 0;
        std::uint32_t pinned     = 0;      // entries written
        // Slots on the shared invalid-opcode handler: landing on one means the previous length
        // was wrong. Unlike `len[i] == 0`, which only means a real length is unknown.
        bool          invalid[kMax] = {};
        std::uint32_t invalidCount  = 0;

        // Lengths pinned from the handler alone that no compiled code has exercised.
        bool          unverified[kMax] = {};
        std::uint32_t unverifiedCount  = 0;

        // Operand-dependent lengths: the count word at +2 adds `varUnit` bytes per unit to `len`.
        // 0 = fixed length.
        std::uint8_t  varUnit[kMax] = {};
        // Slots whose handler addresses R14 and so can name a parameter; without this gate a
        // two-byte opcode's "operand" would be the next instruction's bytes.
        bool          framesR14[kMax] = {};
        std::uint32_t framedCount     = 0;
        std::uint32_t scanMicros      = 0;   // what the arm-time handler scan cost
        bool          ok         = false;
    };

    // Fills the table from the pinned constants, bounded by the verified SlotSet's size.
    bool PinPcodeLengths(const Image& img, const SlotSet& slots,
                         PcodeLengths& out);

    // The lengths for the module armed in this process, or nullptr.
    const PcodeLengths* ArmedLengths();
    void SetArmedLengths(const PcodeLengths& l);
    void ClearArmedLengths();
    // Latched at every arm from XRAYXL_DIAG, whether or not lengths get pinned.
    void SetPcodeDiagnostics(bool on);

    // Argument types for one procedure, indexed by frame slot: argument n is at [R14 + 8n], and
    // a ByVal Variant occupies three slots. Null where unrecoverable, usually an unread parameter.
    struct ArgTypes
    {
        // Tracks vbaargs' kMaxSlots: 60 `ByVal Variant` parameters occupy 181 slots.
        static constexpr int kMax = 192;
        const char*   name[kMax] = {};   // nullptr where unknown
        // The opcode seen at the slot, named or not, so a missing type says which opcode to add.
        std::uint16_t op[kMax]   = {};

        // The exit opcode, which gives the declared return type (ExitReturnKind); the argument
        // walk needs it because a Variant Function gets its result VARIANT in an argument slot.
        // 0 when no exit was reached.
        std::uint16_t exitOp     = 0;

        // For a class or form Function, the offset of the trailing result slot, so "the last
        // slot is not a parameter" is checked against the bytecode.
        std::int32_t  exitOperand = 0;

        // Stopped early or resynchronised, so a `?` may be a skipped load, not an unread parameter.
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

    // The (opcode, operand) stream of one procedure whose parameters did not all resolve, to
    // tell a gap in the tables from one in the walk. Empty when everything resolved.
    const char*   PcodeUntypedDump();

    // The opcodes that stopped walks, most frequent first: an opcode that never occurs costs
    // nothing however unlengthed.
    const char*   PcodeStopLine();

    // The same ranking as a warning, or "" when no walk was stopped.
    const char*   PcodeStopWarning();

    // Opcodes whose length was used and after which the walk broke: a wrong length, worse than
    // a missing one. Carries the clean-walk fraction.
    const char*   PcodeSuspectWarning();
    // Walks whose last statement did not end at ProcSize, by exit; "" when there were none.
    const char*   PcodeClosureWarning();
    // Lengths this session used that nothing has confirmed, ranked; usually "".
    const char*   PcodeUnverifiedWarning();

    // Procedures walked, and those walked cleanly to an exit without a resync. A correct table
    // walks every one cleanly.
    void          PcodeHealth(long long& walks, long long& clean);

    // What the R14 frame gate refused; empty when it refused nothing.
    const char*   PcodeNotFramedWarning();

    // Write every distinct procedure walked this session as raw bytes, for the
    // offline length solver. XRAYXL_DIAG only; call at disarm, never in a hook.
    std::string   WritePcodeCorpus(const std::wstring& path);

    // Stores to the result ([R14-8], or [R14-0x18] for a Variant), whose opcode carries the
    // return type. Scanned, not walked, so it does not depend on the length table. Distinct
    // (opcode, operand) pairs in first-seen order, at most `max`.
    int ReturnStoreCandidates(std::uint64_t trailer, std::uint16_t* ops,
                              std::int32_t* offs, int max);

    // The declared type a store opcode carries (store = load + 32), or nullptr rather than a guess.
    const char* PcodeStoreTypeName(std::uint32_t storeOp);

    // Walks the procedure behind `trailer`. `maxArg` is the argument slot count from argSz; an
    // index above it is discarded. 0 when unknown.
    bool ReadArgTypes(std::uint64_t trailer, ArgTypes& out, int maxArg = 0);
}
