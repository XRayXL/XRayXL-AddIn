#include "vbapcode.h"
#include "core/clock.h"
#include "vbatrailer.h"
#include "vbapcode_tables.h"
#include <windows.h>
#include <cstdio>
#include <cstring>
#include "core/safemem.h"
#include "third_party/minhook/src/hde/hde64.h"
#include "core/log.h"
#include "core/tracemodes.h"
#include <string>

namespace vba
{
    namespace
    {
        bool CorpusWalked(std::uint32_t slot)
        {
            const std::uint32_t byteIdx = slot >> 3;
            if (byteIdx >= sizeof kCorpusWalked) return false;
            return (kCorpusWalked[byteIdx] & (1u << (slot & 7))) != 0;
        }

        // A ranking truncated too short hides most of its own finding, and the
        // whole point of the line is to be worked through. Disarm-time only.
        constexpr int kTopStops = 40;


        PcodeLengths  g_armed;
        bool          g_haveArmed = false;

        volatile LONG64 g_declines[static_cast<int>(PcDecline::Count_)] = {};
        volatile LONG64 g_walks = 0;
        volatile LONG64 g_resyncs = 0;

        // WHICH opcode stopped a walk, per slot. The denominator that matters
        // is not "how many of the 1023 implemented slots have a length" --
        // nothing cares about slots that never occur -- but "which opcodes cost
        // us a signature, and how often". That distribution is head-heavy.
        volatile LONG g_stopOp[PcodeLengths::kMax] = {};
        // The opcode whose length was just used when the walk then broke. `g_stopOp` names an
        // opcode with no length; this names one whose length is probably wrong, which is worse:
        // a missing length resynchronises safely, a wrong one steps into the middle of an
        // operand.
        volatile LONG g_suspectOp[PcodeLengths::kMax] = {};

        // The same blame, but certain: the walk landed on a slot that is not an instruction, so
        // the step that brought it there was wrong. A bisect starts here.
        volatile LONG g_definiteOp[PcodeLengths::kMax] = {};

        // REFUSED BY THE FRAME GATE: an argument-shaped operand on an opcode
        // whose handler never addresses R14, so it cannot be naming a parameter.
        // A high count on one opcode means the gate is working; a high count on
        // a NAMED load opcode would mean the frame scan is wrong.
        volatile LONG g_notFramedOp[PcodeLengths::kMax] = {};

        // DEVELOPER ONLY. With XRAYXL_DIAG the dump also captures a walk that
        // RESYNCHRONISED rather than only one that failed to type a parameter --
        // different questions, and only the first shows where a walk went wrong.
        // Read once at arm (class F), so a user's session cannot reach the copy.
        bool g_dumpResync = false;

        // Steps taken through a length nothing has confirmed. Bounded by the
        // same array the other rankings use, and only touched for the slots the
        // bitmap marks -- zero cost on every ordinary opcode.
        volatile LONG g_unverifiedOp[PcodeLengths::kMax] = {};

        // Walked from offset 0 to a clean end with no resynchronisation. THE
        // health number for the length table: a correct table walks every
        // procedure cleanly, so anything less is a defect and this sizes it.
        volatile LONG64 g_cleanWalks = 0;

        // ONE UNTYPED PROCEDURE'S (opcode, operand) STREAM -- the instrument
        // that says whether an unrecovered type is a gap in the TABLES or in the
        // WALK, which a `(?)` cannot. Bounded, once per arm.
        constexpr int kDumpMax = 64;
        volatile LONG g_dumpTaken = 0;
        volatile LONG g_dumpWriting = 0;   // one dump writer at a time
        int           g_dumpN = 0;
        std::uint16_t g_dumpOp[kDumpMax] = {};
        std::int32_t  g_dumpOperand[kDumpMax] = {};
        std::uint64_t g_dumpTrailer = 0;

        // THE RAW BYTES, because the decoded stream is only as good as the
        // lengths that produced it: once a walk desynchronises, what it prints
        // is already fiction. The bytes are the only account that does not depend
        // on the thing under investigation.
        constexpr int kRawMax = 192;
        int           g_rawN = 0;
        std::uint8_t  g_raw[kRawMax] = {};

        // The corpus: every distinct procedure this session walked. A length table is correct
        // exactly when every procedure parses from offset 0 to its exit consuming exactly
        // ProcSize, so the bytes are kept and solved offline.
        //
        // DIAG only, bounded, deduplicated by trailer. Claimed lock-free on the hot path and
        // written at disarm, never from a traced thread.
        constexpr int kCorpusProcs = 512;
        // 2048, not 256: a procedure over the cap is DROPPED, not truncated, and
        // at 256 that dropped five of six drivers written to emit the VCall and
        // ImpAdCall families -- so the corpus could never contain the very
        // opcodes it was collected to measure. DIAG-only memory: 512 x 2 KB.
        constexpr int kCorpusBytes = 2048;
        struct CorpusProc
        {
            std::uint64_t trailer  = 0;
            std::uint32_t procSize = 0;
            std::uint16_t n        = 0;                 // bytes actually kept
            std::uint8_t  raw[kCorpusBytes] = {};
        };
        CorpusProc    g_corpus[kCorpusProcs];
        volatile LONG g_corpusN       = 0;
        volatile LONG g_corpusSeen    = 0;   // distinct procedures offered
        volatile LONG g_corpusDropped = 0;   // ...and lost to the cap or the size
        bool          g_corpusOn      = false;

        void NoteDecline(PcDecline d)
        {
            InterlockedIncrement64(&g_declines[static_cast<int>(d)]);
        }

        // The top `max` entries of a per-slot counter, highest first, above
        // `floor`; ties rank the lower slot first. `consume` zeroes what it
        // takes, so a ranking is reported once per arm. Disarm-time only.
        int TopSlots(volatile LONG* counts, long floor, int max, bool consume,
                     int* slot, long* val)
        {
            long snap[PcodeLengths::kMax];
            for (int op = 0; op < PcodeLengths::kMax; ++op) snap[op] = counts[op];
            int n = 0;
            for (; n < max; ++n)
            {
                int at = -1; long best = floor;
                for (int op = 0; op < PcodeLengths::kMax; ++op)
                    if (snap[op] > best) { best = snap[op]; at = op; }
                if (at < 0) break;
                slot[n] = at; val[n] = best; snap[at] = 0;
                if (consume) InterlockedExchange(&counts[at], 0);
            }
            return n;
        }

        // Appends " op<slot>=<count>" per entry at `j` and returns the new
        // length. The shape every log parser reads.
        int AppendOps(char* b, size_t cap, int j, const int* slot, const long* val, int n)
        {
            for (int k = 0; k < n; ++k)
            {
                const int w = _snprintf_s(b + j, cap - j, _TRUNCATE, " op%d=%ld", slot[k], val[k]);
                if (w < 0) break;
                j += w;
            }
            return j;
        }

        // core/safemem.h.
        using core::RdU16;
        using core::RdI32;

        // No length derivation here: the lengths are the constant kSigLength, and deriving them
        // is offline work. An unknown opcode is flagged by number where it stops a walk.
    }

    // Opcodes that touch a parameter slot and carry no type. `?opNNN` would wrongly say "add a
    // row to kTypeOps" for these, so they render a bare `?`.
    //
    //    671  `movsxd rax,[rsi]; add rax,r14; push rax` pushes the ADDRESS of the frame slot
    //         without dereferencing it, which says nothing about what the slot holds.
    //
    // Not here: 739, which does load a value, so `?op739` is the right thing to say.
    bool PcodeCarriesNoType(std::uint32_t op)
    {
        // 662 shares 671's handler. 1122, 1467 and 1470 are the same idiom without the push.
        // 751 forwards a by-reference slot's address, alike for a `ByRef Long` and a `ByRef
        // Double()`.
        return op == 671 || op == 662 || op == 751 ||
               op == 1122 || op == 1467 || op == 1470;
    }

    const char* PcodeTypeName(std::uint32_t op)
    {
        for (const TypeOp& t : kTypeOps) if (t.op == op) return t.name;
        return nullptr;
    }

    // A store names its type through the load it mirrors: store = load + 32. Exceptions, ByVal
    // side only: the String store is 708 (not 663+32) and the Variant store is 694 (not
    // 1477+32). nullptr for those; callers keep explicit entries.
    const char* PcodeStoreTypeName(std::uint32_t storeOp)
    {
        if (storeOp < 32) return nullptr;
        // The subtraction must exclude the exceptions too, or 1509 answers "Variant" through
        // 1477 and 695 "String" through 663.
        if (storeOp == 1477 + 32 || storeOp == 663 + 32) return nullptr;
        // The ByRef Variant store has three forms, picked by the right-hand side:
        // a number 774, `Set` 783, a full copy 787. Only 774 mirrors a load
        // (742); 783 and 787 sit over 751 and 755, which name no type.
        //
        if (storeOp == 783 || storeOp == 787) return "Variant&";
        return PcodeTypeName(storeOp - 32);
    }

    // Two kinds of stop. An opcode that stopped several walks is a real instruction missing
    // from kSigLength, and is actionable. Distinct opcodes each stopping once are operand bytes
    // read as opcodes after the walk lost alignment. A fuzz run asserts only that the repeat
    // group is empty.
    const char* PcodeStopWarning()
    {
        static char b[2048];
        long total = 0, repeats = 0, singles = 0;
        for (int op = 0; op < PcodeLengths::kMax; ++op)
        {
            const long c = g_stopOp[op];
            if (c <= 0) continue;
            total += c;
            if (c >= 2) ++repeats; else ++singles;
        }
        if (total <= 0) { b[0] = 0; return b; }

        int j = _snprintf_s(b, _TRUNCATE,
                    "VBA p-code: %ld walk stop(s) at opcodes of UNKNOWN LENGTH -- each "
                    "resynchronises to the next statement and SKIPS what lies between, "
                    "so a parameter loaded there reads '?' and its signature is marked "
                    "'~'.", total);
        if (repeats)
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                    " ACTIONABLE: %ld opcode(s) stopped more than one walk -- real "
                    "instructions missing from kSigLength. Check each length against "
                    "the corpus consensus, then add it.", repeats);
        else
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                    " No opcode stopped more than one walk, so none is clearly a table "
                    "gap; %ld singleton(s) are most likely operand bytes read after a "
                    "desync.", singles);
        _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, " Ranking: %s", PcodeStopLine());
        return b;
    }

    // What the frame gate refused: opcodes with an argument-shaped operand whose handler never
    // computes an address from R14. A named load opcode here means the frame scan is wrong and
    // is suppressing real parameter types; an unnamed one is the gate working. At DEBUG.
    const char* PcodeNotFramedWarning()
    {
        static char b[512];
        b[0] = 0;
        long total = 0;
        for (int op = 0; op < PcodeLengths::kMax; ++op) total += g_notFramedOp[op];
        if (!total) return b;
        int j = _snprintf_s(b, _TRUNCATE,
                    "VBA p-code: frame gate refused %ld argument-shaped operand(s) whose "
                    "opcode cannot address the frame:", total);
        int slot[8]; long val[8];
        const int n = TopSlots(g_notFramedOp, 0, 8, true, slot, val);
        for (int k = 0; k < n; ++k)
        {
            const char* nm = PcodeTypeName(static_cast<std::uint32_t>(slot[k]));
            const int w = _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, " op%d=%ld%s",
                                      slot[k], val[k], nm ? " (NAMED LOAD -- SCAN SUSPECT)" : "");
            if (w < 0) break;
            j += w;
        }
        _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, ".");
        return b;
    }

    // Which lengths are probably wrong: opcodes whose length was used and after which the walk
    // could not continue. Repeat offenders only, because a single break after an unrelated
    // desync blames whatever opcode preceded it.
    const char* PcodeSuspectWarning()
    {
        static char b[1400];
        const long long walks = g_walks, cleanN = g_cleanWalks;
        if (walks <= 0) { b[0] = 0; return b; }
        // ALWAYS SAID, suspects or not: reporting it only alongside something
        // wrong would hide the number exactly when it is good news worth
        // trusting.
        int j = _snprintf_s(b, _TRUNCATE,
                    "VBA p-code: %lld of %lld procedure(s) walked cleanly (offset 0 to a "
                    "clean end, no resynchronisation)", cleanN, walks);
        // Definite first: a landing on the invalid-opcode handler proves the previous length
        // wrong, where a stop only suggests it. Both rankings are consumed here, so they are
        // reported once per arm.
        int slot[8]; long val[8];
        const int nd = TopSlots(g_definiteOp, 0, 6, true, slot, val);
        if (nd)
        {
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                    ". LENGTHS PROVEN WRONG (an unresynchronised walk used them and "
                    "landed on a slot that is not an instruction):");
            j = AppendOps(b, sizeof b, j, slot, val, nd);
        }
        const int ns = TopSlots(g_suspectOp, 1, 8, true, slot, val);
        if (ns)
        {
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, nd
                    ? ".  ALSO SUSPECT:"
                    : ". LENGTHS THAT MAY BE WRONG (an unresynchronised walk used them "
                      "and then could not continue):");
            j = AppendOps(b, sizeof b, j, slot, val, ns);
        }
        if (!nd && !ns)
        {
            _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                        ". No length is blamed: every break followed a resynchronisation, "
                        "so the walk was already lost and the remaining cost is opcodes "
                        "with NO length, not wrong ones.");
            return b;
        }
        _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                    ". Bisecting settles each: the length that raises clean walks "
                    "WITHOUT raising desyncs is the right one -- reading the handler "
                    "cannot settle it for a CALL. Please report these with the log.");
        return b;
    }

    const char* PcodeUnverifiedWarning()
    {
        static char b[900];
        b[0] = 0;
        int slot[10]; long val[10];
        const int n = TopSlots(g_unverifiedOp, 0, 10, true, slot, val);
        if (!n) return b;
        int j = _snprintf_s(b, _TRUNCATE,
                    "VBA p-code: this session used length(s) that NO RUNNING CODE HAS "
                    "EVER CONFIRMED -- pinned from the handler alone, never walked "
                    "past in the offline corpus. The walk did not break on them, so the "
                    "trace stands; they are now the best candidates to measure, because "
                    "this workbook writes a shape no generator does. ACTIONABLE: capture "
                    "with XRAYXL_DIAG=1 and send the p-code corpus it writes:");
        j = AppendOps(b, sizeof b, j, slot, val, n);
        _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, ".");
        return b;
    }

    void PcodeHealth(long long& walks, long long& clean)
    {
        walks = g_walks;
        clean = g_cleanWalks;
    }

    const char* PcDeclineName(PcDecline d)
    {
        switch (d)
        {
        case PcDecline::NoLengths:      return "no derived length table";
        case PcDecline::NoProcSize:     return "ProcSize unreadable";
        case PcDecline::ProcSizeInsane: return "ProcSize out of range";
        case PcDecline::CodeUnreadable: return "p-code unreadable";
        case PcDecline::Desynced:       return "walk desynced";
        case PcDecline::UnknownOpcode:  return "stopped at an opcode of unknown length";
        default:                        return "?";
        }
    }
    void ResetPcodeCounts()
    {
        for (int i = 0; i < static_cast<int>(PcDecline::Count_); ++i) g_declines[i] = 0;
        g_walks = 0;
        g_resyncs = 0;
        for (int i = 0; i < PcodeLengths::kMax; ++i)
        { g_stopOp[i] = 0; g_suspectOp[i] = 0; g_definiteOp[i] = 0; g_notFramedOp[i] = 0;
          g_unverifiedOp[i] = 0; }
        g_cleanWalks = 0;
        InterlockedExchange(&g_dumpTaken, 0);
        InterlockedExchange(&g_dumpWriting, 0);
        g_dumpN = 0; g_dumpTrailer = 0; g_rawN = 0;
        InterlockedExchange(&g_corpusN, 0);
        InterlockedExchange(&g_corpusSeen, 0);
        InterlockedExchange(&g_corpusDropped, 0);
        for (int i = 0; i < kCorpusProcs; ++i) g_corpus[i].trailer = 0;
    }

    // Does this handler ever compute an address from R14? R14 is the VBA frame base and
    // argument n is at `[R14 + 8n]`, so a handler that never names it cannot touch a parameter.
    //
    // R14 is encoding 6 with the matching REX bit and can arrive four ways: as modrm.reg with
    // REX.R, as modrm.rm with REX.B, or inside a SIB as index with REX.X or base with REX.B. A
    // false negative silently loses a parameter type.
    //
    // Unconditional jumps are followed, because the frame access is often in a shared tail. The
    // next-opcode fetch, a return or an indirect jump ends the scan.
    bool HandlerTouchesR14(const Image& img, std::uint32_t rva)
    {
        for (int step = 0; step < 400; ++step)
        {
            // Read, not Scan: a mapped page of the image can still fault.
            std::uint8_t p[16];
            if (!img.Read(rva, p, sizeof p)) return false;
            hde64s hs{};
            const unsigned n = hde64_disasm(p, &hs);
            if (!n || (hs.flags & F_ERROR)) return false;
            if (hs.flags & F_MODRM)
            {
                if (hs.modrm_reg == 6 && hs.rex_r) return true;
                if (hs.modrm_rm  == 6 && hs.rex_b) return true;
                if (hs.modrm_rm  == 4 && hs.modrm_mod != 3)   // SIB present
                {
                    if (hs.sib_index == 6 && hs.rex_x) return true;
                    if (hs.sib_base  == 6 && hs.rex_b) return true;
                }
            }
            if (hs.opcode == 0xE9) { rva += n + static_cast<std::uint32_t>(static_cast<std::int32_t>(hs.imm.imm32)); continue; }
            if (hs.opcode == 0xEB) { rva += n + static_cast<std::uint32_t>(static_cast<std::int8_t>(hs.imm.imm8));   continue; }
            if (hs.opcode == 0xFF && (hs.flags & F_MODRM) && hs.modrm_reg == 4) return false;
            if (hs.opcode == 0xC3 || hs.opcode == 0xC2) return false;
            if (hs.opcode == 0x0F && (hs.opcode2 == 0xB7 || hs.opcode2 == 0xBF)
                && hs.modrm_rm == 6 && !hs.rex_b && hs.modrm_mod != 3) return false;
            rva += n;
        }
        return false;
    }

    // Nothing is derived: the lengths ARE the table. The image and SlotSet are
    // still taken because the caller has verified them, and because a slot beyond
    // this build's table must not be written.
    bool PinPcodeLengths(const Image& img, const SlotSet& s, PcodeLengths& out)
    {
        out = PcodeLengths{};
        if (!s.found || !s.verified || s.slots == 0 || s.slots > PcodeLengths::kMax)
            return false;
        out.slots = s.slots;

        for (const SigLength& e : kSigLength)
        {
            // A slot past the end of THIS build's table is not ours to write.
            // The caller checks the slot count, so this fires only if that ever
            // changes -- and then declines rather than writing off the end.
            if (e.slot >= out.slots) continue;
            out.len[e.slot] = e.len;
            ++out.pinned;
        }
        for (const VarLength& e : kVarLength)
        {
            if (e.slot >= out.slots) continue;
            out.len[e.slot]     = e.base;
            out.varUnit[e.slot] = e.unit;
            ++out.pinned;
        }
        // PINNED BUT NEVER EXERCISED. Marked here rather than listed by hand so
        // the set shrinks automatically as the corpus grows: regenerate
        // kCorpusWalked and whatever the new procedures walked past stops being
        // reported.
        for (std::uint32_t i = 0; i < out.slots && i < PcodeLengths::kMax; ++i)
            if (out.len[i] && !CorpusWalked(i)) { out.unverified[i] = true; ++out.unverifiedCount; }
        // Which slots are not instructions: those pointing at the shared invalid-opcode
        // handler. Read once at arm into a bitmap, so the hot path costs an index. The same
        // pass answers which slots can name a parameter (framesR14).
        {
            // Timed because arming cost matters and this pass decodes every handler.
            const long long t0 = core::QpcMicros();
            for (std::uint32_t i = 0; i < out.slots; ++i)
            {
                std::uint64_t va = 0;
                if (!img.Read(s.tableRva + i * 8u, &va, sizeof va)) continue;
                if (va < img.Base()) continue;
                const std::uint32_t rva = static_cast<std::uint32_t>(va - img.Base());
                if (s.invalidHandlerRva && rva == s.invalidHandlerRva)
                { out.invalid[i] = true; ++out.invalidCount; continue; }
                if (HandlerTouchesR14(img, rva))
                { out.framesR14[i] = true; ++out.framedCount; }
            }
            out.scanMicros = static_cast<std::uint32_t>(core::QpcMicros() - t0);
        }
        // Fail open if the scan found nothing, fail closed per slot otherwise. A build whose
        // handlers this decoder cannot follow leaves framedCount zero, and the walk must go on
        // attributing or every parameter type disappears. A low non-zero count is the dangerous
        // middle: it suppresses most attribution while looking like it worked.
        if (out.framedCount && out.framedCount < 64)
        {
            char w[192];
            _snprintf_s(w, sizeof w, _TRUNCATE,
                        "VBA: frame-addressing scan found only %u slot(s) of %u -- "
                        "too few to trust; parameter attribution left ungated",
                        out.framedCount, out.slots);
            core::Log::Warning(w);
            for (std::uint32_t i = 0; i < out.slots; ++i) out.framesR14[i] = false;
            out.framedCount = 0;
        }
        out.ok = (out.pinned > 0);
        return out.ok;
    }

    const PcodeLengths* ArmedLengths() { return g_haveArmed ? &g_armed : nullptr; }
    void SetArmedLengths(const PcodeLengths& l) { g_armed = l; g_haveArmed = l.ok; }
    void ClearArmedLengths() { g_haveArmed = false; }
    void SetPcodeDiagnostics(bool on) { g_dumpResync = on; g_corpusOn = on; }

    namespace
    {
        // The next instruction boundary after losing alignment. Every statement starts with a
        // beginning-of-statement opcode, so the next one is a known-good boundary. A candidate
        // is accepted only if the instruction after it also decodes. Scans one parity of the
        // byte stream in two-byte steps. Returns the new offset, or -1.
        int ResyncScan(std::uint64_t code, std::uint32_t procSize,
                       std::uint32_t from, const PcodeLengths* L)
        {
            for (std::uint32_t j = from; j + 8 <= procSize; j += 2)
            {
                std::uint16_t cand = 0;
                if (!RdU16(code + j, cand)) return -1;
                if (!IsBosSlot(cand)) continue;

                std::uint16_t next = 0;
                if (!RdU16(code + j + 6, next)) return -1;   // BoS is 6 bytes

                // The next word must be a plausible opcode: inside the dispatch table, and not
                // a second beginning-of-statement. Requiring a known length instead would
                // reject genuine boundaries.
                if (next >= L->slots) continue;
                if (IsBosSlot(next))  continue;
                return static_cast<int>(j);
            }
            return -1;
        }

        // Not every instruction is an even number of bytes: `LitI2_Byte` (1643) is three, so
        // everything after it in that procedure sits at an odd offset. The walk's own parity is
        // scanned first, because it is right whenever the step that got here was; then the
        // other.
        int ResyncAtStatement(std::uint64_t code, std::uint32_t procSize,
                              std::uint32_t from, const PcodeLengths* L)
        {
            const int hit = ResyncScan(code, procSize, from, L);
            if (hit >= 0) return hit;
            if (from + 1 + 8 > procSize) return -1;
            return ResyncScan(code, procSize, from + 1, L);
        }
    }

    // Every candidate store near the result, named or not. Both offsets, because a Variant
    // result lives at [R14-0x18] rather than [R14-8].
    int ReturnStoreCandidates(std::uint64_t trailer, std::uint16_t* ops,
                              std::int32_t* offs, int max)
    {
        if (!trailer || max <= 0) return 0;
        std::uint16_t procSize = 0;
        if (!RdU16(trailer + kTrl_procSize, procSize)) return 0;
        if (procSize < 4 || procSize > 0x4000) return 0;
        const std::uint64_t code = trailer - procSize;
        int n = 0;
        for (std::uint32_t i = 2; i + 4 <= procSize && n < max; ++i)
        {
            std::int32_t operand = 0;
            if (!RdI32(code + i, operand)) break;
            if (operand != -8 && operand != -0x18) continue;
            std::uint16_t op = 0;
            if (!RdU16(code + i - 2, op)) continue;
            bool dup = false;
            for (int k = 0; k < n; ++k)
                if (ops[k] == op && offs[k] == operand) { dup = true; break; }
            if (dup) continue;
            ops[n] = op; offs[n] = operand; ++n;
        }
        return n;
    }

    namespace
    {
        // Linear scan for the duplicate check: the table is small, this is
        // DIAG-only, and a hash would need a collision policy that could
        // silently drop a distinct procedure.
        // DIAG only, so a lock is affordable; without it two threads can record one procedure twice.
        SRWLOCK g_corpusLock = SRWLOCK_INIT;
        struct CorpusLock
        {
            CorpusLock()  { AcquireSRWLockExclusive(&g_corpusLock); }
            ~CorpusLock() { ReleaseSRWLockExclusive(&g_corpusLock); }
        };

        void CorpusOffer(std::uint64_t trailer, std::uint64_t code, std::uint32_t procSize)
        {
            const CorpusLock held;
            const LONG n = InterlockedCompareExchange(&g_corpusN, 0, 0);
            for (LONG i = 0; i < n && i < kCorpusProcs; ++i)
                if (g_corpus[i].trailer == trailer) return;      // already have it

            // DROPPED rather than truncated: a truncated body cannot be told
            // from one that ends early, and a solver fed truncated bytes would
            // blame a length for the cut.
            InterlockedIncrement(&g_corpusSeen);
            if (procSize == 0 || procSize > kCorpusBytes)
            { InterlockedIncrement(&g_corpusDropped); return; }

            const LONG slot = InterlockedIncrement(&g_corpusN) - 1;
            if (slot >= kCorpusProcs)
            { InterlockedDecrement(&g_corpusN); InterlockedIncrement(&g_corpusDropped); return; }

            CorpusProc& c = g_corpus[slot];
            int got = 0;
            for (; got < static_cast<int>(procSize) && got < kCorpusBytes; ++got)
            {
                std::uint16_t b = 0;
                if (!RdU16(code + got, b)) break;    // guarded, one byte at a time
                c.raw[got] = static_cast<std::uint8_t>(b & 0xFF);
            }
            c.procSize = procSize;
            c.n        = static_cast<std::uint16_t>(got);
            // Written LAST, so a reader that sees a non-zero trailer sees a
            // complete record rather than a half-filled slot.
            c.trailer  = trailer;
        }

        // What one instruction says about a parameter slot. [R14 + 8n] is argument n, and every
        // typed load has a 4-byte operand at +2.
        //
        // Gated on the frame set, not the operand's shape: a two-byte opcode has no operand, so
        // the bytes at +2 belong to the next instruction. The gate fails open when the arm-time
        // scan found nothing.
        //
        // First evidence wins. A load names the type directly; a write-only parameter emits a
        // typed store, and store = load + 32. With neither, the opcode is kept as a diagnostic.
        // Control-flow opcodes are excluded: a statement marker's offset can look like a frame
        // offset.
        void Attribute(ArgTypes& out, const PcodeLengths* L, std::uint16_t op,
                       std::int32_t operand, bool haveOperand, int maxArg)
        {
            if (!haveOperand || operand <= 0 || (operand % 8) != 0) return;
            const bool framed = (L->framedCount == 0) || L->framesR14[op];
            if (!framed) { InterlockedIncrement(&g_notFramedOp[op]); return; }

            const int idx = operand / 8;
            // Bounded by the frame: an index past argSz came from a bad
            // resync landing, not from a load.
            const int cap = (maxArg > 0 && maxArg < ArgTypes::kMax) ? maxArg + 1
                                                                   : ArgTypes::kMax;
            if (idx < 1 || idx >= cap || out.name[idx]) return;

            const bool  control = IsBosSlot(op) || IsExitSlot(op);
            const char* tn = PcodeTypeName(op);
            if (!tn && !control) tn = PcodeStoreTypeName(op);
            if (tn)                            { out.name[idx] = tn; out.op[idx] = op; }
            else if (!control && !out.op[idx])   out.op[idx] = op;
        }

        // The instructions one walk met, for the untyped dump.
        struct Seen
        {
            std::uint16_t op[kDumpMax];
            std::int32_t  operand[kDumpMax];
            int           n;
            void Note(std::uint16_t o, std::int32_t opd)
            { if (n < kDumpMax) { op[n] = o; operand[n] = opd; ++n; } }
        };

        // Developer evidence after a walk. The corpus takes every procedure. The one-procedure
        // dump takes the first walk with an unnamed argument slot; under XRAYXL_DIAG a walk
        // that resynchronised takes it instead, last wins. A procedure with no argument slots
        // never qualifies.
        void OfferEvidence(std::uint64_t trailer, std::uint64_t code, std::uint32_t procSize,
                           const ArgTypes& out, int maxArg, int resyncs, const Seen& seen)
        {
            if (g_corpusOn) CorpusOffer(trailer, code, procSize);

            bool untyped = false;
            const int top = (maxArg < ArgTypes::kMax) ? maxArg : ArgTypes::kMax - 1;
            for (int q = 1; q <= top && !untyped; ++q) untyped = (out.name[q] == nullptr);
            const bool take =
                (untyped && InterlockedCompareExchange(&g_dumpTaken, 1, 0) == 0) ||
                (g_dumpResync && resyncs > 0 && (InterlockedExchange(&g_dumpTaken, 1), true));
            if (!take) return;
            // The resync path takes the dump without a CAS, so writers exclude each other here.
            if (InterlockedCompareExchange(&g_dumpWriting, 1, 0) != 0) return;

            g_dumpTrailer = trailer;
            // The whole body, up to the cap -- guarded, one byte at a time.
            const int want = (procSize < kRawMax) ? static_cast<int>(procSize) : kRawMax;
            int got = 0;
            for (; got < want; ++got)
            {
                std::uint16_t b = 0;
                if (!RdU16(code + got, b)) break;
                g_raw[got] = static_cast<std::uint8_t>(b & 0xFF);
            }
            g_rawN  = got;
            g_dumpN = seen.n;
            for (int q = 0; q < seen.n; ++q)
            { g_dumpOp[q] = seen.op[q]; g_dumpOperand[q] = seen.operand[q]; }
            InterlockedExchange(&g_dumpWriting, 0);
        }
    }

    bool ReadArgTypes(std::uint64_t trailer, ArgTypes& out, int maxArg)
    {
        out = ArgTypes{};
        const PcodeLengths* L = ArmedLengths();
        if (!L) { NoteDecline(PcDecline::NoLengths); return false; }
        if (!trailer) return false;

        std::uint16_t procSize = 0;
        if (!RdU16(trailer + kTrl_procSize, procSize))
        { NoteDecline(PcDecline::NoProcSize); return false; }
        // A body of 0 bytes is not a body, and the ceiling is generous but
        // finite so a stale trailer cannot send the walk into unrelated memory.
        if (procSize < 2 || procSize > 0x4000)
        { NoteDecline(PcDecline::ProcSizeInsane); return false; }

        const std::uint64_t code = trailer - procSize;
        std::uint32_t i = 0;
        bool clean = false;
        // Where this statement says the next one starts; 0 if it is the last.
        std::uint32_t stmtNext = 0;
        int  resyncs = 0;
        constexpr int kMaxResyncs = 64;   // bounded: a procedure is finite
        // The opcode whose length was last used to step: when the walk cannot
        // continue, THIS is the suspect, not the byte it landed on.
        std::uint16_t prevOp = 0;
        Seen seen{};

        // Alignment is lost: an opcode outside the table, a landing on the invalid handler, or
        // an opcode with no pinned length. Keep what was recovered and resynchronise on the
        // next statement marker. No guessed step is taken over an unsized opcode, since a wrong
        // step attributes types to the wrong parameters.
        //
        // `blame` records prevOp only from an unresynchronised run; after a resync the fault
        // may be several steps back. Returns false when nothing recognisable remains.
        auto lose = [&](PcDecline why, volatile LONG* blame) -> bool
        {
            NoteDecline(why);
            if (prevOp && resyncs == 0) InterlockedIncrement(&blame[prevOp]);
            if (resyncs >= kMaxResyncs) return false;
            const int j = ResyncAtStatement(code, procSize, i + 2, L);
            if (j < 0) return false;
            i = static_cast<std::uint32_t>(j);
            ++resyncs;
            InterlockedIncrement64(&g_resyncs);
            return true;
        };

        while (i + 2 <= procSize)
        {
            std::uint16_t op = 0;
            if (!RdU16(code + i, op)) { NoteDecline(PcDecline::CodeUnreadable); return false; }

            // Outside the table, prevOp's length is a suspect. On the invalid
            // handler (none of its slots is an operation) the step was
            // certainly wrong: a definite suspect, and the landing is a stop.
            if (op >= L->slots)
            {
                if (lose(PcDecline::Desynced, g_suspectOp)) continue;
                break;
            }
            if (L->invalidCount && L->invalid[op])
            {
                InterlockedIncrement(&g_stopOp[op]);
                if (lose(PcDecline::UnknownOpcode, g_definiteOp)) continue;
                break;
            }

            // An `Exit` emits the same opcode as the real end. Which it is comes
            // from the statement carrying it: a statement's operand is the next
            // statement's offset, and is ZERO on the last one. So an exit needs no
            // length -- and there are ten of them, one per return type.
            //
            if (IsProcTerminatorSlot(op))
            {
                clean = true;
                if (!out.exitOp)
                {
                    out.exitOp = op;
                    std::int32_t exOperand = 0;
                    if (RdI32(code + i + 2, exOperand)) out.exitOperand = exOperand;
                }
                if (stmtNext <= i) break;       // last statement: the real end
                prevOp = op;
                i = stmtNext;
                continue;           // a terminator is not a parameter load
            }

            // The type is recorded before the length is consulted, so an
            // unsized opcode still names its slot.
            std::int32_t operand = 0;
            const bool haveOperand = RdI32(code + i + 2, operand);
            // A statement opcode says where the next statement begins, relative
            // to itself. Only forward offsets inside the procedure are kept: a
            // backward or out-of-range one is not a next statement.
            if (IsBosSlot(op))
            {
                // Reset on every statement, or the last one inherits a successor
                // and its exit looks early.
                stmtNext = (haveOperand && operand > 0 &&
                            i + static_cast<std::uint32_t>(operand) <= procSize)
                         ? i + static_cast<std::uint32_t>(operand) : 0;
            }
            seen.Note(op, haveOperand ? operand : 0);
            Attribute(out, L, op, operand, haveOperand, maxArg);

            std::uint32_t len = L->len[op];
            if (len && L->varUnit[op])
            {
                std::uint16_t count = 0;
                if (!RdU16(code + i + 2, count)) { NoteDecline(PcDecline::CodeUnreadable); return false; }
                len += static_cast<std::uint32_t>(L->varUnit[op]) * count;
            }
            // No pinned length: either a real instruction missing from
            // kSigLength, or prevOp's length is wrong and `op` is operand
            // debris (how 1309 presented, as a phantom "opcode 2"). Both are
            // recorded; only evidence outside this walk separates them.
            if (len == 0)
            {
                InterlockedIncrement(&g_stopOp[op]);
                if (lose(PcDecline::UnknownOpcode, g_suspectOp)) continue;
                break;
            }

            // One array index; the increment fires only for a length no
            // running code has confirmed.
            if (L->unverified[op]) InterlockedIncrement(&g_unverifiedOp[op]);
            prevOp = op;
            i += len;
        }

        if (i == procSize) clean = true;
        if (clean && resyncs == 0) InterlockedIncrement64(&g_cleanWalks);
        InterlockedIncrement64(&g_walks);
        // Stopped early or resynchronised: a `?` may be a skipped load.
        out.partial = !clean || resyncs > 0;
        OfferEvidence(trailer, code, procSize, out, maxArg, resyncs, seen);
        return true;
    }

    const char* PcodeStopLine()
    {
        static char b[1600];
        int slot[kTopStops]; long val[kTopStops];
        const int n = TopSlots(g_stopOp, 0, kTopStops, false, slot, val);
        const int j = _snprintf_s(b, _TRUNCATE, "VBA p-code stops:%s", n ? "" : " none");
        AppendOps(b, sizeof b, j, slot, val, n);
        return b;
    }

    // One procedure's instruction stream; empty when every procedure this
    // session was fully typed. See the globals for why it is bounded.
    const char* PcodeUntypedDump()
    {
        static char line[1400];
        line[0] = 0;
        if (InterlockedCompareExchange(&g_dumpTaken, 0, 0) == 0 || g_dumpN == 0)
            return line;
        int j = _snprintf_s(line, sizeof(line), _TRUNCATE,
                            "VBA p-code untyped proc 0x%llX:",
                            static_cast<unsigned long long>(g_dumpTrailer));
        if (j < 0) { line[0] = 0; return line; }
        for (int q = 0; q < g_dumpN && j < static_cast<int>(sizeof(line)) - 24; ++q)
        {
            const int w = _snprintf_s(line + j, sizeof(line) - j, _TRUNCATE,
                                      " %u@%d", static_cast<unsigned>(g_dumpOp[q]),
                                      static_cast<int>(g_dumpOperand[q]));
            if (w < 0) break;
            j += w;
        }
        if (g_rawN > 0 && j < static_cast<int>(sizeof(line)) - 16)
        {
            const int w = _snprintf_s(line + j, sizeof(line) - j, _TRUNCATE, " | raw:");
            if (w > 0) j += w;
            for (int q = 0; q < g_rawN && j < static_cast<int>(sizeof(line)) - 6; ++q)
            {
                const int w2 = _snprintf_s(line + j, sizeof(line) - j, _TRUNCATE,
                                           "%s%02X", (q && (q % 2) == 0) ? " " : "",
                                           static_cast<unsigned>(g_raw[q]));
                if (w2 < 0) break;
                j += w2;
            }
        }
        return line;
    }

    // Writes the corpus, at disarm and never from a hook. One line per procedure: trailer,
    // ProcSize, bytes. No decoded opcodes or lengths, because the file exists to check the
    // lengths. The summary includes what was dropped, so a corpus missing every large procedure
    // does not flatter the table.
    std::string WritePcodeCorpus(const std::wstring& path)
    {
        char msg[256];
        const LONG n = InterlockedCompareExchange(&g_corpusN, 0, 0);
        const LONG kept = (n < kCorpusProcs) ? n : kCorpusProcs;
        if (!g_corpusOn)
            return "VBA p-code corpus: not collected (set XRAYXL_DIAG=1 before starting Excel)";
        if (kept <= 0)
            return "VBA p-code corpus: nothing walked, nothing written";

        FILE* f = nullptr;
        if (_wfopen_s(&f, path.c_str(), L"a") != 0 || !f)
        {
            _snprintf_s(msg, sizeof msg, _TRUNCATE,
                        "VBA p-code corpus: %ld procedure(s) held but the file could not be opened", kept);
            return msg;
        }
        // APPEND, not truncate: the reset at arm clears the in-memory set, so
        // truncating here would keep only the LAST arm. Duplicates across arms
        // are the solver's problem and trivial to remove; losing procedures is
        // not recoverable.
        fprintf(f, "# XRayXL p-code corpus. One procedure per line: raw bytes only.\n");
        fprintf(f, "# A length table is correct exactly when every one of these parses\n");
        fprintf(f, "# from offset 0 to an exit opcode, consuming exactly size bytes.\n");
        int written = 0;
        for (LONG i = 0; i < kept; ++i)
        {
            const CorpusProc& c = g_corpus[i];
            if (!c.trailer || !c.n) continue;          // slot never completed
            fprintf(f, "proc %llX size=%u bytes=",
                    static_cast<unsigned long long>(c.trailer), c.procSize);
            for (int q = 0; q < c.n; ++q) fprintf(f, "%02X", static_cast<unsigned>(c.raw[q]));
            fprintf(f, "\n");
            ++written;
        }
        fclose(f);
        _snprintf_s(msg, sizeof msg, _TRUNCATE,
                    "VBA p-code corpus: %d procedure(s) written (%ld offered, %ld dropped "
                    "as too large or over the cap)",
                    written, InterlockedCompareExchange(&g_corpusSeen, 0, 0),
                    InterlockedCompareExchange(&g_corpusDropped, 0, 0));
        return msg;
    }

    const char* PcodeLine()
    {
        static char b[400];
        const PcodeLengths* L = ArmedLengths();
        int j = _snprintf_s(b, _TRUNCATE, "VBA p-code: %s, %llu walks",
                            L ? "lengths pinned" : "NO length table",
                            static_cast<unsigned long long>(g_walks));
        if (L)
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                             " (%u opcodes pinned of %u slots, %u non-instruction"
                             " slots known, %llu resyncs)",
                             L->pinned, L->slots, L->invalidCount,
                             static_cast<unsigned long long>(g_resyncs));
        for (int i = 0; i < static_cast<int>(PcDecline::Count_); ++i)
        {
            if (!g_declines[i]) continue;
            const int w = _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, ", %s=%llu",
                                      PcDeclineName(static_cast<PcDecline>(i)),
                                      static_cast<unsigned long long>(g_declines[i]));
            if (w < 0) break;
            j += w;
        }
        return b;
    }
}
