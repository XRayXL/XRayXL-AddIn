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

        // THE OPCODE WHOSE LENGTH WE JUST USED, when the walk then broke. Where
        // `g_stopOp` names an opcode with NO length, this names one whose length
        // is probably WRONG -- the worse defect, because a missing length
        // resynchronises safely (and is marked `~`) while a wrong one steps into
        // the middle of an operand and reads what is there as an instruction.
        //
        // The rule: if stepping by `len[prev]` lands somewhere the walk cannot
        // continue, the suspect is `prev`, not where we landed. It is the only
        // in-process signal separating the two, and it names the opcode a bisect
        // should try.
        volatile LONG g_suspectOp[PcodeLengths::kMax] = {};

        // THE SAME BLAME, BUT CERTAIN: the walk landed on a slot that is NOT AN
        // INSTRUCTION (677 of the 1700 share the invalid handler), so the step
        // that brought it there was wrong. `g_suspectOp` is only a heuristic, so
        // the two are ranked apart and a bisect starts here.
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

        // ---- THE CORPUS. Every DISTINCT procedure this session walked. ----
        //
        // The single dump above answers "show me one procedure that went
        // wrong"; it cannot be pointed at a real workbook and asked WHICH
        // lengths are wrong.
        //
        // A LENGTH TABLE IS CORRECT EXACTLY WHEN EVERY PROCEDURE PARSES FROM
        // OFFSET 0 TO ITS EXIT, CONSUMING EXACTLY ProcSize -- checkable against
        // any VBA that exists, needing no external truth, and not depending on a
        // generator happening to emit the right construct (which the shape fuzzer
        // demonstrably does not: seed 4242 was byte-identical either side of a
        // real fix). So keep the BYTES of everything walked and solve offline.
        //
        // DIAG only, bounded, deduplicated by trailer -- a procedure called in a
        // loop is one observation, not a thousand. Claimed lock-free on the hot
        // path and WRITTEN AT DISARM, never from a traced thread.
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

        // NO LENGTH DERIVATION HERE: the lengths are the measured constant in
        // kSigLength above; deriving them is offline work, which is where a question
        // about a NEW VBE7 build belongs. An unknown opcode is flagged by number
        // at the point it stops a walk.
    }

    // OPCODES THAT TOUCH A PARAMETER SLOT AND CARRY NO TYPE -- known, not
    // missing. A `?opNNN` tells a reader "add a row to kTypeOps"; for these that
    // advice is wrong, so the honest rendering is a bare `?`.
    //
    //   671  `movsxd rax,[rsi]; add rax,r14; push rax` computes the ADDRESS of
    //        the frame slot and pushes it WITHOUT dereferencing -- an address-of,
    //        how an array assignment names its destination. That a slot's address
    //        was taken says nothing about what the slot holds.
    //
    // Deliberately NOT here: 739, which DOES load a value. It carries a type; it
    // simply matches no VBA declared scalar, so if it ever appears `?op739` is
    // the right thing to say. [measured: the handlers for 671, 739, 747]
    bool PcodeCarriesNoType(std::uint32_t op)
    {
        // 662 SHARES 671'S HANDLER EXACTLY. Naming a TYPE from a shared handler
        // is forbidden, but this is not a type claim -- it is the
        // observation that the operation conveys NO type, which the shared
        // handler proves for both slots at once.
        //
        // 1122, 1467 and 1470 are the same idiom without the push: compute
        // `r14 + operand` and fetch the next opcode, leaving the address for the
        // instruction that follows.
        //
        // All four render a bare `?` rather than `?opNNN` -- "there is no type"
        // rather than "a table entry is missing", which is the difference between
        // a fact and a wrong instruction to the reader.
        // [measured: the handlers for 662, 671, 1122, 1467, 1470]
        return op == 671 || op == 662 || op == 1122 || op == 1467 || op == 1470;
    }

    const char* PcodeTypeName(std::uint32_t op)
    {
        for (const TypeOp& t : kTypeOps) if (t.op == op) return t.name;
        return nullptr;
    }

    // A STORE NAMES ITS TYPE THROUGH THE LOAD IT MIRRORS: store = load + 32.
    // Measured on seven ByVal types and five ByRef positions, and it
    // also explains 770, the ByRef Long store.
    //
    // A RELATION RATHER THAN MORE ROWS, because each row would otherwise be found
    // only when somebody happened to write a procedure that used it.
    //
    // KNOWN EXCEPTIONS, ByVal side only: the String store is 708 (not 663+32) and
    // the Variant store is 694 (not 1477+32). nullptr for those rather than a
    // wrong name; callers keep explicit entries.
    const char* PcodeStoreTypeName(std::uint32_t storeOp)
    {
        if (storeOp < 32) return nullptr;
        // THE RELATION RUNS BOTH WAYS AND ONLY ONE DIRECTION IS TRUE. The
        // exceptions above say the ByVal Variant store is 694, not 1477+32 --
        // so the SUBTRACTION must exclude them too, or 1509 answers "Variant"
        // through 1477. op1509's handler touches R14 on neither measured build,
        // so it cannot store to a frame slot at all: that answer was a confident
        // wrong type on a parameter, not a diagnostic. 663+32 = 695 is the same
        // case, harmless today only because 663 is absent from the load table.
        if (storeOp == 1477 + 32 || storeOp == 663 + 32) return nullptr;
        return PcodeTypeName(storeOp - 32);
    }

    // TWO KINDS OF STOP, AND ONLY ONE IS ACTIONABLE. An opcode that stopped
    // SEVERAL walks is being met repeatedly at real instruction boundaries, so
    // it is an instruction and its absence from kSigLength is a gap worth
    // closing. A scatter of DISTINCT opcodes each stopping ONCE is the opposite:
    // operand bytes read as opcodes after the walk lost alignment.
    //
    // Both are reported and only the first is called actionable. A fuzz run
    // asserts the REPEAT group is empty; asserting on the singletons would make
    // the oracle cry wolf.
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

    // WHAT THE FRAME GATE REFUSED: opcodes met with an argument-shaped operand
    // -- positive, a multiple of eight, inside the frame's argument count --
    // whose handler never computes an address from R14.
    //
    // Two readings, opposite in meaning:
    //   a NAMED load opcode here  -> the frame scan is wrong and is suppressing
    //                                real parameter types. A defect, and loud.
    //   an unnamed opcode here    -> the gate caught what it exists to catch.
    // At DEBUG, because on a healthy build it is noise about noise.
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

    // WHICH LENGTHS ARE PROBABLY WRONG, and how healthy the table is. Separate
    // from PcodeStopWarning, which names opcodes with NO length: this names ones
    // whose length was USED and after which the walk could not continue. Repeat
    // offenders only, because a single break after an unrelated desync blames
    // whatever opcode happened to precede it.
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
        // Definite first: a landing on the invalid-opcode handler PROVES the
        // previous length wrong where a mere stop only suggests it, and ranking
        // them together buries the certain evidence under the guesses. Suspects
        // are repeat offenders only, because a single break after an unrelated
        // desync blames whatever opcode happened to precede it. Both rankings
        // are consumed here, so they are reported once per arm.
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

    // DOES THIS HANDLER EVER COMPUTE AN ADDRESS FROM R14? R14 is the VBA frame
    // base and argument n is at `[R14 + 8n]`, so a handler that never names it
    // cannot touch a parameter whatever its operand looks like. Checked against
    // a known answer from the handlers: all 42 named load opcodes in
    // kTypeOps come out framed, 42/42, on both measured builds.
    //
    // R14 is encoding 6 with the matching REX bit and can arrive FOUR ways -- as
    // modrm.reg with REX.R, as modrm.rm with REX.B, or inside a SIB as index with
    // REX.X or base with REX.B. Missing any one clears a handler that does use
    // the frame, and this predicate SUPPRESSES attribution, so a false negative
    // silently loses a parameter type.
    //
    // Unconditional jumps are followed, because a handler's frame access is often
    // in the shared tail it jumps to. The next-opcode fetch ends the
    // instruction's own work; a return or indirect jump ends the handler.
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
        // WHICH SLOTS ARE NOT INSTRUCTIONS: 677 of the 1700 point at one shared
        // invalid-opcode handler, and a walk that LANDS on one is already lost
        // rather than meeting an instruction it does not know. Read once at arm
        // into a bitmap -- class T, so the hot path costs an index.
        //
        // The same pass answers which slots can name a parameter (framesR14), so
        // one read serves both and the invalid handler is skipped rather than
        // decoded once per slot that uses it.
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
        // FAIL OPEN IF THE SCAN FOUND NOTHING, FAIL CLOSED PER SLOT OTHERWISE.
        // A build whose handlers this decoder cannot follow leaves `framedCount`
        // zero, and the walk must go on attributing or every parameter
        // type in the session disappears at once. A LOW but non-zero count is the
        // dangerous middle -- it suppresses most attribution while looking like
        // it worked -- and a count near the named load family's size has not
        // understood this build.
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
        // The next instruction boundary after losing alignment. Every VBA
        // statement starts with a beginning-of-statement opcode, so the next one
        // is a known-good boundary -- and whatever threw us sits INSIDE a
        // statement, not across one. A candidate is accepted only if the
        // instruction after it also decodes, which is what stops a BoS-shaped
        // pair of bytes inside an operand being mistaken for a real one.
        //
        // Returns the new offset, or -1 if nothing recognisable remains.
        // One parity of the byte stream. Two-byte steps, because instructions
        // are two-byte aligned RELATIVE TO EACH OTHER for all but one opcode --
        // see the caller for the one.
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

                // The next word must be a PLAUSIBLE OPCODE -- inside the
                // dispatch table, and not a second beginning-of-statement, which
                // no real statement starts with. Requiring a known LENGTH
                // instead rejects genuine boundaries; the walk's argument-index
                // bound is what rejects a bad landing.
                if (next >= L->slots) continue;
                if (IsBosSlot(next))  continue;
                return static_cast<int>(j);
            }
            return -1;
        }

        // NOT EVERY INSTRUCTION IS AN EVEN NUMBER OF BYTES. `LitI2_Byte` (1643)
        // carries a single signed byte, so it is THREE bytes and everything
        // after it in that procedure sits at an odd offset. Scanning in
        // two-byte steps from where the walk got lost therefore searches the
        // parity the walk HAD, which is the right one only while the walk was
        // synchronised -- and resynchronising is what we do when it was not.
        // With the wrong parity every real boundary is invisible, the scan
        // returns -1, and the whole procedure is abandoned instead of partly
        // recovered.
        //
        // So: the walk's own parity first, because it is right whenever the
        // step that got here was, and it costs a session nothing to prefer it;
        // then the other. A scan that finds nothing is the only case that pays
        // for the second pass. [measured: handler 1643 advances RSI by 3]
        //
        // WHAT THE SECOND PASS IS WORTH, measured by planting an odd length
        // (`LitI2` 4 -> 5) on an opcode this compiler does emit, so the walk
        // is driven onto odd offsets. Same workbook, same plant: with one
        // parity the walk desynced ONCE and stopped dead -- 0 resyncs -- while
        // with both it resynchronised 15 times and kept recovering. The types
        // recovered were the same 3 incomplete signatures either way, so this
        // buys continuation, not (on that sample) extra answers.
        // [measured: family probe, builds A/B]
        // THE SECOND SCAN IS THE ODD PARITY. Stepping two bytes searches the
        // parity the walk HAD, which is the right one only while the walk is
        // synchronised -- and resynchronising is what happens when it is not.
        int ResyncAtStatement(std::uint64_t code, std::uint32_t procSize,
                              std::uint32_t from, const PcodeLengths* L)
        {
            const int hit = ResyncScan(code, procSize, from, L);
            if (hit >= 0) return hit;
            if (from + 1 + 8 > procSize) return -1;
            return ResyncScan(code, procSize, from + 1, L);
        }
    }

    // EVERY candidate store near the result, named or not, for MEASUREMENT: the
    // seven load/store pairs cover the numeric types, and String, Variant and
    // Object returns store through opcodes the pairing was never measured on. Which
    // opcodes they DO use is a question about the bytecode, and this asks it.
    //
    // BOTH OFFSETS, because a Variant result lives at [R14-0x18] rather than
    // [R14-8]: a scan looking only for -8 would report "no store" for a procedure
    // that has one, making "absent" and "not looked for" the same answer again.
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

        // What one instruction says about a parameter slot. [R14 + 8n] is
        // argument n, and every typed load has a 4-byte operand at +2.
        //
        // Gated on the frame set, not on the operand's shape: a two-byte opcode
        // has no operand, so the four bytes at +2 belong to the NEXT
        // instruction, and "positive, a multiple of eight" proves nothing.
        // Refusals are counted. The gate fails open when the arm-time scan
        // answered for nobody, or a build the scan cannot follow would lose
        // every type at once.
        //
        // First evidence wins. A load names the type directly; a write-only
        // parameter emits no load but a typed store, and store = load + 32
        // (the relation the return decoder uses), which is how `p = 9.75`
        // recovers `Double`. With neither, the opcode is kept as a diagnostic
        // so a table gap names itself. Control-flow opcodes are excluded: a
        // statement marker's offset can look exactly like a frame offset.
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

        // Developer evidence after a walk. The corpus takes every procedure:
        // failures prove a defect, successes constrain the candidate lengths.
        // The one-procedure dump takes the first walk with an unnamed argument
        // slot; under XRAYXL_DIAG a walk that resynchronised takes it instead
        // (last wins), since a resync is the earlier symptom and a driver that
        // resyncs on entry must not hide the procedure under investigation.
        // A procedure with no argument slots never qualifies.
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
        int  resyncs = 0;
        constexpr int kMaxResyncs = 64;   // bounded: a procedure is finite
        // The opcode whose length was last used to step: when the walk cannot
        // continue, THIS is the suspect, not the byte it landed on.
        std::uint16_t prevOp = 0;
        Seen seen{};

        // Alignment is lost -- an opcode outside the table, a landing on the
        // invalid handler, or an opcode with no pinned length -- and the
        // recovery is the same: keep what was recovered (it was synchronised)
        // and resynchronise on the next statement marker. No guessed step is
        // ever taken over an unsized opcode; a wrong step attributes types to
        // the wrong parameters. The stated cost: a write-only String assigned
        // a literal loses its type (slot 11 has no length on any of 42 builds).
        //
        // `blame` records prevOp -- its length made the step that got here --
        // only from an unresynchronised run; after a resync the fault may be
        // several steps back. Returns false when nothing recognisable remains.
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

            // An exit ends the procedure cleanly. Which exit names the return
            // type and its operand is the class/form result slot; the argument
            // walk needs both.
            if (IsProcTerminatorSlot(op))
            {
                clean = true;
                out.exitOp = op;
                std::int32_t exOperand = 0;
                if (RdI32(code + i + 2, exOperand)) out.exitOperand = exOperand;
                break;
            }

            // The type is recorded before the length is consulted, so an
            // unsized opcode still names its slot.
            std::int32_t operand = 0;
            const bool haveOperand = RdI32(code + i + 2, operand);
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

    // WRITE THE CORPUS. At DISARM, never from a hook: this opens a file.
    //
    // One line per procedure -- trailer, ProcSize, bytes. Everything the offline
    // solver needs and nothing it has to trust: no decoded opcodes, no lengths,
    // no interpretation, because the point is to CHECK the lengths and a file
    // that had applied them would be circular.
    //
    // The summary includes what was DROPPED: a corpus silently omitting every
    // large procedure would make the table look better than it is, and large
    // procedures are where a wrong length has most room to go unnoticed.
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
