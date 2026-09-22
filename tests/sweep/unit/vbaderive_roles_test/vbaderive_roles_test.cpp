// UNIT TEST for vba::Derive (vba/vbaderive.cpp): no handler may be patched in
// two roles. The patcher emits one stub per ORIGINAL handler, so an exit slot
// holding the beginning-of-statement handler would share whichever thunk was
// emitted first.
//
// A synthetic image: a 1700-slot table of distinct code pointers with the exit
// repeat-groups and the BoS pair made uniform. Needs no Excel and no VBE7.
//
// Built by XRayXL.sln into build\x64\Release\unit\.

#include "vbaderive.h"
#include "vbapcode_tables.h"

#include <cstdio>
#include <cstring>
#include <vector>

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    constexpr std::uint64_t kBase   = 0x10000000ull;
    constexpr std::uint32_t kTable  = 0x1000;
    constexpr std::uint32_t kSlots  = 1700;
    constexpr std::uint32_t kCodeLo = 0x100000;
    constexpr std::uint32_t kCodeHi = 0x200000;

    class FakeImage : public vba::Image
    {
    public:
        FakeImage() : m_bytes(kTable + kSlots * 8 + 0x1000, 0)
        {
            for (std::uint32_t i = 0; i < kSlots; ++i) Set(i, Handler(i));
            Set(1039, Handler(1038)); Set(1040, Handler(1038));                      // exit group of three
            for (std::uint32_t s : { 630u, 631u, 634u, 635u }) Set(s, Handler(628)); // exit group of five
            Set(1645, Handler(615));                                                  // the BoS pair
            Set(1646, Handler(616));                                                  // its breakpoint twins
        }
        static std::uint64_t Handler(std::uint32_t i) { return kBase + kCodeLo + 0x10ull * i; }
        void Set(std::uint32_t slot, std::uint64_t v) { std::memcpy(&m_bytes[kTable + slot * 8], &v, 8); }

        bool Read(std::uint32_t rva, void* dst, std::size_t n) const override
        {
            if (rva > m_bytes.size() || n > m_bytes.size() - rva) return false;
            std::memcpy(dst, &m_bytes[rva], n);
            return true;
        }
        const std::uint8_t* Scan(std::uint32_t rva, std::size_t n) const override
        {
            if (rva > m_bytes.size() || n > m_bytes.size() - rva) return nullptr;
            return &m_bytes[rva];
        }
        std::uint64_t Base() const override       { return kBase; }
        std::uint32_t SizeOfImage() const override { return static_cast<std::uint32_t>(m_bytes.size()); }
        std::uint32_t CodeLoRva() const override   { return kCodeLo; }
        std::uint32_t CodeHiRva() const override   { return kCodeHi; }
        bool          Ok() const override          { return true; }

    private:
        std::vector<std::uint8_t> m_bytes;
    };
}

int main()
{
    // ---- the control: a well-formed table verifies --------------------------
    {
        FakeImage img;
        const vba::SlotSet s = vba::Derive(img);
        printf("  control: %s\n", vba::Describe(s).c_str());
        Check(s.found && s.verified, "a well-formed table verifies");
        Check(s.endOk, "the End slot verifies on it");
        Check(s.bosBpOk, "the BosBp pair verifies on it");
        int bp = 0;
        for (const vba::PatchSite& site : s.sites)
            if (std::strcmp(site.role, "bosbp") == 0 && (site.slot == 616 || site.slot == 1646)) ++bp;
        Check(bp == 2, "both BosBp slots are patch sites");
        Check(vba::IsBosSlot(616) && vba::IsBosSlot(1646), "the walk takes BosBp as a statement");
    }

    // ---- the BosBp pair disagrees: breakpoints go unhandled, the arm does not --
    {
        FakeImage img;
        img.Set(1646, FakeImage::Handler(1646));
        const vba::SlotSet s = vba::Derive(img);
        Check(!s.bosBpOk, "a BosBp pair holding two handlers does not verify");
        Check(s.verified, "and the table still verifies");
    }

    // ---- the BosBp pair holds BoS's handler: not a distinct role ---------------
    {
        FakeImage img;
        img.Set(616, FakeImage::Handler(615)); img.Set(1646, FakeImage::Handler(615));
        const vba::SlotSet s = vba::Derive(img);
        Check(!s.bosBpOk, "a BosBp pair holding the BoS handler does not verify");
    }

    // ---- one exit slot holds the BoS handler ---------------------------------
    {
        FakeImage img;
        img.Set(0x2088 / 8, FakeImage::Handler(615));
        const vba::SlotSet s = vba::Derive(img);
        printf("  exit slot = BoS handler: %s\n", vba::Describe(s).c_str());
        Check(s.found, "the table is still found");
        Check(!s.verified, "an exit slot holding the BoS handler refuses to verify");
    }

    // ---- a whole group holds it: uniform, so only the role check catches it ---
    {
        FakeImage img;
        for (std::uint32_t sl : { 1038u, 1039u, 1040u }) img.Set(sl, FakeImage::Handler(615));
        const vba::SlotSet s = vba::Derive(img);
        printf("  exit group = BoS handler: %s\n", vba::Describe(s).c_str());
        Check(!s.verified, "an exit group holding the BoS handler refuses to verify");
    }

    // ---- an exit that ends its statement has no length ------------------------
    // The walk never steps over one, so a length there could be wrong and nothing would notice.
    {
        int withLength = 0;
        for (const vba::SigLength& e : vba::kSigLength)
            if (vba::IsProcTerminatorSlot(e.slot)) { printf("  exit %u has length %u\n", e.slot, e.len); ++withLength; }
        Check(withLength == 0, "no terminating exit slot carries a length");

        // ...its length is in kExitLength instead, for the closure check; 953 is unmeasured.
        int unmeasured = 0, notExit = 0;
        for (std::uint32_t s = 0; s < 1700; ++s)
        {
            if (!vba::IsProcTerminatorSlot(s) || s == 953) continue;
            bool have = false;
            for (const vba::SigLength& e : vba::kExitLength) have = have || e.slot == s;
            if (!have) { printf("  exit %u has no measured length\n", s); ++unmeasured; }
        }
        for (const vba::SigLength& e : vba::kExitLength)
            if (!vba::IsProcTerminatorSlot(e.slot)) { printf("  %u is not an exit\n", e.slot); ++notExit; }
        Check(unmeasured == 0, "every terminating exit but 953 has a measured length");
        Check(notExit == 0, "kExitLength lists only terminating exits");
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
