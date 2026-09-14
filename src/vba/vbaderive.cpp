#include "vbaderive.h"
#include "core/safemem.h"
#include "core/hash.h"
#include "vbaslots.h"

#include <windows.h>
#include <algorithm>
#include <sstream>
#include <cstring>

namespace vba
{
    namespace
    {
        // ---------------------------------------------------------------
        // Slot indices. NOT addresses -- see the header for why these are a
        // legitimate constant and a derived address is not. Byte offsets from
        // the prior art divided by 8; recorded here in slot form because that
        // is the unit that is portable.
        // ---------------------------------------------------------------
        constexpr std::uint32_t kBosSlot[2] = { 0x1338 / 8, 0x3368 / 8 };   // 615, 1645

        constexpr std::uint32_t kExitSlot[] = {
            0x2088/8, 0x2070/8, 0x2078/8, 0x2080/8, 0x2090/8,
            0x1DC0/8, 0x3400/8, 0x1DC8/8,
            0x13A0/8, 0x13B0/8, 0x13B8/8, 0x13D0/8, 0x13D8/8,
            0x2EB0/8, 0x3408/8, 0x2EB8/8,
            0x0FC0/8, 0x1380/8, 0x1388/8, 0x1390/8, 0x1398/8,
            0x1378/8, 0x1360/8, 0x33F0/8, 0x33F8/8,
        };
        // How the exit slots group into runs that must share one handler; only multi-slot groups can fail.
        constexpr int kExitGroup[] = { 1,3,1,1,1,1,5,1,1,1,1,1,1,1,1,1,1,1,1 };

        // THE RAISE OPCODE, which makes error attribution possible. 497 is the
        // one that actually RAISES and fires four times per Err.Raise
        // [measured]; two neighbours are decoys -- 718 fires when a
        // handler is merely SET UP, 1272 on an Err object access.
        //
        // ITS PORTABILITY IS CHECKED, not assumed: the BoS pair and the exit
        // slots are corpus-grade and 497 never was. Its fingerprint is handler
        // SHARING -- 497 holds a handler of its own on every build measured,
        // while its decoys sit in a five-slot group.
        // [measured]
        constexpr std::uint32_t kRaiseSlot = 0xF88 / 8;   // 497

        // THE `End` OPCODE, which tears the whole VBA session down. Microsoft's
        // own PDB names slot 619 `lblEX_End`, between `lblEX_Debug` (618) and the
        // GoSub `lblEX_Return` (620) this project already pins -- so the two
        // slots either side of it are independently confirmed by what the tracer
        // does with them today.
        //
        // ITS PORTABILITY IS INHERITED, NOT ASSUMED. The fingerprint is the same
        // one the raise slot uses -- 619 holds a handler of its OWN -- and being
        // a singleton is EXACTLY what the equivalence partition records. That
        // partition is byte-identical across the corpus and is already verified
        // at arm time (`kPartitionHash`), so a build where 619 stopped being a
        // singleton could not match the hash. The check below is still made
        // rather than argued, because a fingerprint that is never read is not a
        // check.
        constexpr std::uint32_t kEndSlot = 619;

        constexpr std::uint32_t kExpectedSlots = 1700;

        // THE OPCODE SET WE MEASURED. FNV-1a 64 over the dispatch table's
        // equivalence partition (for every slot, the lowest slot sharing its
        // handler). Byte-identical on 7.1.10.33 and 7.1.11.58, whose handler
        // addresses agree nowhere. A mismatch does not mean the table is the
        // wrong run -- the structural checks answer that -- it means the opcode
        // SET is not the one `kSigLength` describes, so the lengths do not apply.
        // [measured]
        constexpr std::uint64_t kPartitionHash = 0x077934407C79A117ULL;
        constexpr std::uint32_t kMinRun        = 256;   // a run shorter than this is noise

        static_assert(sizeof(kExitSlot) / sizeof(kExitSlot[0]) == 25, "25 exit slots");
        constexpr int ExitGroupTotal() { int n = 0; for (int g : kExitGroup) n += g; return n; }
        static_assert(ExitGroupTotal() == 25, "the exit groups cover exactly the 25 exit slots");

        bool GuardedQword(const std::uint8_t* p, std::uint64_t& out)
        {
            __try
            {
                out = static_cast<std::uint64_t>(p[0])
                    | (static_cast<std::uint64_t>(p[1]) << 8)
                    | (static_cast<std::uint64_t>(p[2]) << 16)
                    | (static_cast<std::uint64_t>(p[3]) << 24)
                    | (static_cast<std::uint64_t>(p[4]) << 32)
                    | (static_cast<std::uint64_t>(p[5]) << 40)
                    | (static_cast<std::uint64_t>(p[6]) << 48)
                    | (static_cast<std::uint64_t>(p[7]) << 56);
                return true;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
        }

    }

    const char* DeclineName(Decline d)
    {
        switch (d)
        {
        case Decline::ImageUnreadable:        return "a page of the module faulted while scanning";
        case Decline::ModuleNotLoaded:        return "VBE7 not loaded";
        case Decline::HeaderUnreadable:       return "PE header unreadable";
        case Decline::NotPe64:                return "not a 64-bit PE";
        case Decline::NoExecutableSection:    return "no executable section";
        case Decline::NoCandidateRun:         return "no pointer run long enough";
        case Decline::NoCandidateAgreedOnBos: return "no candidate table agreed on the BoS pair";
        case Decline::WrongSlotCount:         return "table is not 1700 slots";
        case Decline::ExitGroupNotUniform:    return "an exit repeat-group held more than one handler";
        case Decline::BosSlotsDiffer:         return "the two BoS slots differ";
        case Decline::ExitSharesBosHandler:   return "an exit slot holds the beginning-of-statement handler";
        default:                              return "?";
        }
    }

    // -------------------------------------------------------------------
    // The derivation
    // -------------------------------------------------------------------
    bool IsBosSlot(std::uint32_t slot)
    {
        for (std::uint32_t s : kBosSlot) if (s == slot) return true;
        return false;
    }

    bool IsExitSlot(std::uint32_t slot)
    {
        for (std::uint32_t s : kExitSlot) if (s == slot) return true;
        return false;
    }

    // Exit-family slots that fire mid-procedure, so the type walk must not stop there:
    // GoSub `Return` and the pre-exit cleanups do not end a procedure (vbaslots.h).
    // Keep in step with ProcedureEnd in vbaboundary.cpp.
    bool IsProcTerminatorSlot(std::uint32_t slot)
    {
        if (slot == kSlot_ZeroRetVal || slot == kSlot_ZeroRetValVar ||
            slot == kSlot_GoSubReturn) return false;
        return IsExitSlot(slot);
    }

    // A class or form Function returns its value through a TRAILING argument
    // slot -- the COM `[out, retval]` convention -- and leaves through
    // ExitProcCbHresult (0x3400/8) or ExitProcFrameCbHresult (0x3408/8), whose
    // own operand is the byte offset of that slot. Measured: a class
    // `Function(ByVal a As Long, Optional b As Variant) As Long` ends
    // `1664 ExitProcCbHresult operand=24`, and 24 is slot 3 of a 4-slot frame
    // whose parameters are slots 1 and 2. A class Sub leaves through
    // ExitProcHresult and has no such slot.
    bool ExitHasTrailingResultSlot(std::uint32_t slot)
    {
        return slot == kSlot_ExitProcCbHresult || slot == kSlot_ExitProcFrameCbHresult;
    }

    // -------------------------------------------------------------------
    // Derivation, in the order the checks run. Everything after the pick can
    // only REFUSE; nothing reaches back to retry or fall back.
    // -------------------------------------------------------------------
    namespace
    {
        void Note(SlotSet& s, Decline d) { s.declines[static_cast<int>(d)]++; }

        bool ReadSlot(const Image& img, std::uint32_t tableRva, std::uint32_t slot, std::uint64_t& out)
        {
            return img.Read(tableRva + slot * 8, &out, 8);
        }

        // Runs of qwords that all point into executable code: the table's SHAPE,
        // which survives a build where every address in it does not.
        struct Run { std::uint32_t rva, len; };

        // ---- 1. THE SHAPE: every run of >= kMinRun code pointers, longest
        // first. `runnerUp` is the longest rival's length, the margin the
        // report prints. An unreadable page is counted once and skipped whole,
        // rather than taking an exception for every qword in it.
        std::vector<Run> ScanCodePointerRuns(const Image& img, SlotSet& s, std::uint32_t& runnerUp)
        {
            const std::uint64_t base   = img.Base();
            const std::uint64_t codeLo = base + img.CodeLoRva();
            const std::uint64_t codeHi = base + img.CodeHiRva();
            const std::uint32_t end    = img.SizeOfImage();

            std::vector<Run> runs;
            std::uint32_t bestLen = 0, runStart = 0, runLen = 0;
            runnerUp = 0;
            auto closeRun = [&]()
            {
                if (runLen >= kMinRun) runs.push_back({ runStart, runLen });
                if (runLen > bestLen) { runnerUp = bestLen; bestLen = runLen; }
                else if (runLen > runnerUp) runnerUp = runLen;
                runLen = 0;
            };

            bool imageFaulted = false;
            for (std::uint32_t rva = 0; rva + 8 <= end; rva += 8)
            {
                const std::uint8_t* p = img.Scan(rva, 8);
                std::uint64_t v = 0;
                if (p)
                {
                    if (!GuardedQword(p, v))
                    {
                        if (!imageFaulted) { Note(s, Decline::ImageUnreadable); imageFaulted = true; }
                        closeRun();
                        rva = (rva & ~0xFFFu) + 0x1000u - 8u;   // the += 8 lands on the next page
                        continue;
                    }
                }
                else if (!img.Read(rva, &v, 8)) v = 0;

                if (v >= codeLo && v < codeHi)
                {
                    if (runLen == 0) runStart = rva;
                    ++runLen;
                }
                else closeRun();
            }
            closeRun();

            std::sort(runs.begin(), runs.end(),
                      [](const Run& a, const Run& b) { return a.len > b.len; });
            return runs;
        }

        // ---- 2. THE CANDIDATE: the run whose two beginning-of-statement slots
        // hold the same handler. Two loads, no symbol, and it disambiguated
        // correctly on every build in the corpus.
        const Run* PickByBosPair(const Image& img, const std::vector<Run>& runs,
                                 std::uint64_t& bos0, std::uint64_t& bos1)
        {
            for (const Run& r : runs)
            {
                if (r.len <= kBosSlot[1]) continue;                 // too short to hold the pair
                std::uint64_t a = 0, b = 0;
                if (!ReadSlot(img, r.rva, kBosSlot[0], a)) continue;
                if (!ReadSlot(img, r.rva, kBosSlot[1], b)) continue;
                if (a == b) { bos0 = a; bos1 = b; return &r; }
            }
            return nullptr;
        }

        // ---- 3. THE FINGERPRINT. THE EQUIVALENCE PARTITION, hashed: for every
        // slot, the index of the lowest slot holding the same handler. Every
        // address in the table differs between builds and this sequence does
        // not, because it records only which slots AGREE. `kSigLength` is keyed
        // by slot index and was measured against this opcode set, so a build
        // that renumbered the slots would decode every procedure into fiction
        // with nothing to say so. Identical on both measured builds, 847
        // groups. [measured]
        //
        // The same pass counts the distinct handlers for the report -- a table
        // whose slots nearly all coincide would be a run of something else --
        // and finds the MOST FREQUENT handler, the shared invalid-opcode one:
        // no real operation is implemented 600-odd times. Only a genuine
        // majority-of-a-kind is reported as such; a commonest handler covering
        // a handful of slots would not be the invalid handler.
        void FingerprintTable(const Image& img, SlotSet& s)
        {
            std::vector<std::uint64_t> handlers(s.slots, 0);
            for (std::uint32_t i = 0; i < s.slots; ++i) ReadSlot(img, s.tableRva, i, handlers[i]);

            // Lowest slot per handler, found by sorting (handler, slot) rather
            // than by 1700^2 comparisons -- this runs in the user's arm.
            std::vector<std::pair<std::uint64_t, std::uint32_t>> hs;
            hs.reserve(s.slots);
            for (std::uint32_t i = 0; i < s.slots; ++i) hs.push_back({ handlers[i], i });
            std::sort(hs.begin(), hs.end());

            std::vector<std::uint32_t> lowest(s.slots, 0);
            std::uint64_t invalidVa = 0; std::size_t invalidCount = 0; std::uint32_t distinct = 0;
            for (std::size_t i = 0; i < hs.size();)
            {
                std::size_t j = i;
                while (j < hs.size() && hs[j].first == hs[i].first) ++j;
                for (std::size_t k = i; k < j; ++k) lowest[hs[k].second] = hs[i].second;
                if (hs[i].first != 0 && j - i > invalidCount) { invalidCount = j - i; invalidVa = hs[i].first; }
                if (hs[i].first != 0) ++distinct;       // an unreadable slot is not a handler
                i = j;
            }

            std::uint64_t h64 = core::kFnvOffset;
            for (std::uint32_t v : lowest)
                for (int k = 0; k < 4; ++k)
                    h64 = core::Fnv1aByte(h64, static_cast<std::uint8_t>(v >> (k * 8)));
            s.partitionHash = h64;
            // The length table is PINNED, not derived, so arming has to verify
            // it rather than trust it: the partition hash says the slot indices
            // mean what the table thinks.
            s.partitionOk       = (h64 == kPartitionHash);
            s.distinctHandlers  = distinct;
            if (invalidVa && invalidCount >= s.slots / 8)
            {
                s.invalidHandlerRva = static_cast<std::uint32_t>(invalidVa - img.Base());
                s.invalidSlots      = static_cast<std::uint32_t>(invalidCount);
            }
        }

        // ---- 4. THE EXIT SLOTS: every repeat-group must hold exactly one
        // handler. Records each as a patch site; false when any group is not
        // uniform, and that is counted.
        bool VerifyExitGroups(const Image& img, SlotSet& s)
        {
            const std::uint64_t base = img.Base();
            bool ok = true;
            int idx = 0, group = 0;
            for (int gsize : kExitGroup)
            {
                std::uint64_t first = 0;
                bool uniform = true;
                for (int k = 0; k < gsize; ++k, ++idx)
                {
                    std::uint64_t v = 0;
                    if (!ReadSlot(img, s.tableRva, kExitSlot[idx], v)) { uniform = false; continue; }
                    if (k == 0) first = v;
                    else if (v != first) uniform = false;

                    PatchSite site;
                    site.slot       = kExitSlot[idx];
                    site.handlerRva = static_cast<std::uint32_t>(v - base);
                    site.role       = "exit";
                    s.sites.push_back(site);
                }
                if (!uniform) { Note(s, Decline::ExitGroupNotUniform); ok = false; }
                ++group;
            }
            s.exitGroups = group;
            return ok;
        }

        // ---- 5. A SINGLETON SLOT, verified on its own terms.
        //
        // NEITHER the raise slot NOR the End slot is part of `verified`. Error
        // attribution is one feature and `End` handling is another; tracing is
        // the product. A slot that fails its check costs that one feature and
        // nothing else, so it degrades rather than refusing the arm -- and it
        // is reported, because a silently absent feature is the failure this
        // project keeps paying for.
        //
        // THE CHECK IS THE FINGERPRINT THE CORPUS MEASURED: the slot holds a
        // handler of its OWN. For 497 the two ruled-out decoys sit in a shared
        // five-slot group, so a renumbering that slid another opcode into it
        // would almost certainly land on a shared handler and be caught here;
        // for 619 singleton-ness is the very thing `kPartitionHash` pins. And
        // it must not be the BoS handler either -- that would mean the table
        // moved under us in a way the count alone would miss.
        //
        // ONE FUNCTION FOR BOTH, because two copies of an address check is how
        // they drift apart.
        void VerifySingletonSlot(const Image& img, SlotSet& s, std::uint32_t slot,
                                 const char* role, bool& okOut)
        {
            std::uint64_t rv = 0;
            if (slot >= s.slots || !ReadSlot(img, s.tableRva, slot, rv) || rv == 0) return;

            int sharers = 0;
            for (std::uint32_t i = 0; i < s.slots; ++i)
            {
                std::uint64_t v = 0;
                if (ReadSlot(img, s.tableRva, i, v) && v == rv) ++sharers;
            }
            const bool unique = (sharers == 1);
            const bool notBos = (static_cast<std::uint32_t>(rv - img.Base()) != s.bosHandlerRva);
            if (!unique || !notBos) return;

            okOut  = true;
            PatchSite site;
            site.slot       = slot;
            site.handlerRva = static_cast<std::uint32_t>(rv - img.Base());
            site.role       = role;
            s.sites.push_back(site);
        }
    }

    SlotSet Derive(const Image& img)
    {
        SlotSet s;

        if (!img.Ok())
        {
            const Decline why = img.WhyNotOk();
            Note(s, why);
            s.detail = (why == Decline::HeaderUnreadable)
                     ? "VBE7 is loaded but its PE header did not parse"
                     : (why == Decline::NotPe64)
                     ? "VBE7 is loaded but is not a 64-bit PE"
                     : "VBE7 is not loaded in this process";
            return s;
        }
        if (!img.CodeHiRva())     { Note(s, Decline::NoExecutableSection);
                                    s.detail = "no executable section"; return s; }

        // 1. the shape
        const std::vector<Run> runs = ScanCodePointerRuns(img, s, s.runnerUpSlots);
        if (runs.empty()) { Note(s, Decline::NoCandidateRun);
                            s.detail = "no run of >=256 code pointers anywhere in the image"; return s; }

        // 2. the candidate
        std::uint64_t bos0 = 0, bos1 = 0;
        const Run* chosen = PickByBosPair(img, runs, bos0, bos1);
        if (!chosen)
        {
            Note(s, Decline::NoCandidateAgreedOnBos);
            std::ostringstream o;
            o << runs.size() << " candidate run(s), longest " << runs.front().len
              << " slots, none agreed on the BoS pair";
            s.detail = o.str();
            return s;
        }
        s.found         = true;
        s.tableRva      = chosen->rva;
        s.slots         = chosen->len;
        s.bosHandlerRva = static_cast<std::uint32_t>(bos0 - img.Base());

        // ---- verification. Everything below can only REFUSE. ----
        bool ok = true;
        if (bos0 != bos1) { Note(s, Decline::BosSlotsDiffer); ok = false; }   // belt and braces
        if (s.slots != kExpectedSlots) { Note(s, Decline::WrongSlotCount); ok = false; }

        // 3. the fingerprint, 4. the exit groups, then the BoS sites
        FingerprintTable(img, s);
        if (!VerifyExitGroups(img, s)) ok = false;
        for (std::uint32_t bs : kBosSlot)
        {
            PatchSite site;
            site.slot       = bs;
            site.handlerRva = s.bosHandlerRva;
            site.role       = "bos";
            s.sites.push_back(site);
        }
        // One stub per handler, so a handler in two roles would get one role's thunk.
        for (const PatchSite& site : s.sites)
            if (std::strcmp(site.role, "exit") == 0 && site.handlerRva == s.bosHandlerRva)
            { Note(s, Decline::ExitSharesBosHandler); ok = false; break; }

        // 5. the two singleton slots -- features, not the product, so not in `ok`
        VerifySingletonSlot(img, s, kRaiseSlot, "raise", s.raiseOk);
        VerifySingletonSlot(img, s, kEndSlot,   "end",   s.endOk);

        s.verified = ok;
        std::ostringstream o;
        o << (ok ? "verified" : "REFUSED") << ": table at +0x" << std::hex << s.tableRva
          << std::dec << ", " << s.slots << " slots, " << s.distinctHandlers
          << " distinct handlers, " << s.sites.size() << " slots would be patched";
        s.detail = o.str();
        return s;
    }

    std::string Describe(const SlotSet& s)
    {
        std::ostringstream o;
        if (!s.found)
        {
            o << "VBA dispatch table NOT found -- " << s.detail;
        }
        else
        {
            o << "VBA dispatch table +0x" << std::hex << s.tableRva << std::dec
              << "  slots=" << s.slots
              << "  distinct=" << s.distinctHandlers
              << "  runner-up run=" << s.runnerUpSlots
              << "  BoS handler +0x" << std::hex << s.bosHandlerRva << std::dec
              << "  invalid handler +0x" << std::hex << s.invalidHandlerRva << std::dec
              << " x" << s.invalidSlots
              << "  partition=0x" << std::hex << s.partitionHash << std::dec
              << (s.partitionOk ? " (known opcode set)" : " UNKNOWN OPCODE SET")
              << "  exit groups=" << s.exitGroups
              << "  would patch " << s.sites.size() << " slots"
              << "  [" << (s.verified ? "VERIFIED" : "REFUSED -- would not arm") << "]";
        }
        for (int i = 0; i < static_cast<int>(Decline::Count_); ++i)
            if (s.declines[i])
                o << "\n    declined: " << DeclineName(static_cast<Decline>(i))
                  << " x" << s.declines[i];
        return o.str();
    }

    // -------------------------------------------------------------------
    // Images
    // -------------------------------------------------------------------
    namespace
    {
        // The PE header walk. `why` tells an unreadable header from one that is not 64-bit.
        bool ParseHeaders(const std::uint8_t* b, std::size_t n,
                          std::uint32_t& sizeOfImage,
                          std::uint32_t& codeLo, std::uint32_t& codeHi,
                          Decline* why = nullptr)
        {
            const auto fail = [&](Decline d) { if (why) *why = d; return false; };
            if (n < 0x40) return fail(Decline::HeaderUnreadable);
            const auto* dos = reinterpret_cast<const IMAGE_DOS_HEADER*>(b);
            if (dos->e_magic != IMAGE_DOS_SIGNATURE) return fail(Decline::HeaderUnreadable);
            const std::uint32_t pe = static_cast<std::uint32_t>(dos->e_lfanew);
            if (pe + sizeof(IMAGE_NT_HEADERS64) > n) return fail(Decline::HeaderUnreadable);
            const auto* nt = reinterpret_cast<const IMAGE_NT_HEADERS64*>(b + pe);
            if (nt->Signature != IMAGE_NT_SIGNATURE) return fail(Decline::HeaderUnreadable);
            if (nt->OptionalHeader.Magic != IMAGE_NT_OPTIONAL_HDR64_MAGIC)
                return fail(Decline::NotPe64);

            sizeOfImage = nt->OptionalHeader.SizeOfImage;

            const auto* sec = IMAGE_FIRST_SECTION(nt);
            codeLo = 0xFFFFFFFFu; codeHi = 0;
            for (unsigned i = 0; i < nt->FileHeader.NumberOfSections; ++i)
            {
                const auto& sh = sec[i];
                const std::uint32_t va = static_cast<std::uint32_t>(sh.VirtualAddress);
                const std::uint32_t vs = (std::max)(static_cast<std::uint32_t>(sh.Misc.VirtualSize),
                                                    static_cast<std::uint32_t>(sh.SizeOfRawData));
                if (sh.Characteristics & IMAGE_SCN_MEM_EXECUTE)
                {
                    codeLo = (std::min)(codeLo, va);
                    codeHi = (std::max)(codeHi, static_cast<std::uint32_t>(va + vs));
                }
            }
            if (codeLo == 0xFFFFFFFFu) codeLo = 0;
            return true;
        }

        // SEH-only wrappers. Kept free of C++ objects on purpose: __try may not
        // share a frame with anything that requires unwinding.
        bool GuardedParseLoaded(const std::uint8_t* p,
                                std::uint32_t& size, std::uint32_t& lo, std::uint32_t& hi,
                                Decline* why)
        {
            __try { return ParseHeaders(p, 0x1000, size, lo, hi, why); }
            __except (EXCEPTION_EXECUTE_HANDLER)
            {
                // The header itself faulted -- the module is mapped but the
                // first page is not readable. That is "unreadable", not
                // "absent", and the two must not be reported as the same thing.
                if (why) *why = Decline::HeaderUnreadable;
                return false;
            }
        }

    }

    LoadedVbe7::LoadedVbe7()
    {
        HMODULE h = GetModuleHandleW(L"VBE7.DLL");
        if (!h) return;

        m_p = reinterpret_cast<const std::uint8_t*>(h);
        m_base = reinterpret_cast<std::uint64_t>(h);

        // SEH lives in its own function: __try cannot share a frame with
        // objects that need unwinding (C2712), and this one has several.
        m_ok = GuardedParseLoaded(m_p, m_size, m_codeLo, m_codeHi, &m_why);

    }

    bool LoadedVbe7::Read(std::uint32_t rva, void* dst, std::size_t n) const
    {
        if (!m_ok || rva > m_size || n > m_size - rva) return false;
        return core::RdBytes(reinterpret_cast<std::uint64_t>(m_p + rva), dst, n);
    }

    const std::uint8_t* LoadedVbe7::Scan(std::uint32_t rva, std::size_t n) const
    {
        if (!m_ok || rva > m_size || n > m_size - rva) return nullptr;
        return m_p + rva;   // a loaded image is contiguous in RVA space
    }

}
