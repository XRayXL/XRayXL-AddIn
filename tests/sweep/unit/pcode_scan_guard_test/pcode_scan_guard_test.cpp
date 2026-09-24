// Unit test for vba::PinPcodeLengths (vba/vbapcode.cpp): the arm-time scan of
// every dispatch handler must read the image through Image::Read, which fails
// on an unreadable page, and never through a raw Scan pointer, which faults on
// the user's Arm call.
//
// A synthetic image whose table is readable and whose handler page is mapped
// PAGE_NOACCESS -- the shape Decline::ImageUnreadable exists for. The control
// image has a readable handler, `mov rax,[r14+8]; ret`, which must be framed.
//
// Built by XRayXL.sln into build\x64\Release\unit\; the log and the corpus's module walk are
// stubbed below.

#include "vbapcode.h"

#include <windows.h>
#include <cstdio>
#include <cstring>
#include <string>
#include <vector>

namespace core { namespace Log { void Warning(const std::string&) {} } }
namespace vba { int ModuleTrailers(std::uint64_t, std::uint64_t*, int) { return 0; } }   // the corpus is not tested here

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    constexpr std::uint64_t kBase    = 0x10000000ull;
    constexpr std::uint32_t kTable   = 0x1000;
    constexpr std::uint32_t kSlots   = 1700;
    constexpr std::uint32_t kHandler = 0x10000;
    constexpr std::uint32_t kImage   = 0x20000;

    class FakeImage : public vba::Image
    {
    public:
        explicit FakeImage(bool readable) : m_readable(readable), m_table(kSlots * 8)
        {
            const std::uint64_t va = kBase + kHandler;
            for (std::uint32_t i = 0; i < kSlots; ++i) std::memcpy(&m_table[i * 8], &va, 8);
            static const std::uint8_t kCode[] = { 0x49, 0x8B, 0x46, 0x08, 0xC3 };   // mov rax,[r14+8]; ret
            std::memset(m_code, 0xCC, sizeof m_code);
            std::memcpy(m_code, kCode, sizeof kCode);
            m_noAccess = static_cast<std::uint8_t*>(
                VirtualAlloc(nullptr, 0x1000, MEM_COMMIT | MEM_RESERVE, PAGE_NOACCESS));
        }

        bool Read(std::uint32_t rva, void* dst, std::size_t n) const override
        {
            if (rva >= kTable && rva + n <= kTable + m_table.size())
            { std::memcpy(dst, &m_table[rva - kTable], n); return true; }
            if (m_readable && rva >= kHandler && rva + n <= kHandler + sizeof m_code)
            { std::memcpy(dst, m_code + (rva - kHandler), n); return true; }
            return false;
        }
        const std::uint8_t* Scan(std::uint32_t rva, std::size_t n) const override
        {
            if (rva >= kTable && rva + n <= kTable + m_table.size()) return &m_table[rva - kTable];
            if (rva >= kHandler && rva + n <= kHandler + 0x1000)
            {
                if (!m_readable) return m_noAccess + (rva - kHandler);
                return (rva + n <= kHandler + sizeof m_code) ? m_code + (rva - kHandler) : nullptr;
            }
            return nullptr;
        }
        std::uint64_t Base() const override       { return kBase; }
        std::uint32_t SizeOfImage() const override { return kImage; }
        std::uint32_t CodeLoRva() const override   { return kHandler; }
        std::uint32_t CodeHiRva() const override   { return kHandler + 0x1000; }
        bool          Ok() const override          { return true; }

    private:
        bool                      m_readable;
        std::vector<std::uint8_t> m_table;
        std::uint8_t              m_code[64];
        std::uint8_t*             m_noAccess = nullptr;
    };

    // Guarded here so a fault in the code under test reads as a FAIL line.
    bool PinGuarded(const vba::Image& img, const vba::SlotSet& s, vba::PcodeLengths& pl, bool& faulted)
    {
        faulted = false;
        __try { return vba::PinPcodeLengths(img, s, pl); }
        __except (EXCEPTION_EXECUTE_HANDLER) { faulted = true; return false; }
    }
}

int main()
{
    vba::SlotSet s;
    s.found = true; s.verified = true; s.slots = kSlots; s.tableRva = kTable;

    static vba::PcodeLengths pl;
    bool faulted = false;

    // ---- the control: a readable handler that addresses R14 is framed --------
    {
        FakeImage img(true);
        const bool ok = PinGuarded(img, s, pl, faulted);
        Check(!faulted && ok, "control: the lengths pin");
        printf("  control framed %u of %u\n", pl.framedCount, pl.slots);
        Check(pl.framedCount == kSlots, "control: a readable handler addressing R14 is framed");
    }

    // ---- the handler page is mapped but unreadable ----------------------------
    {
        FakeImage img(false);
        const bool ok = PinGuarded(img, s, pl, faulted);
        Check(!faulted, "an unreadable handler page does not fault the scan");
        Check(ok, "the lengths still pin");
        Check(pl.framedCount == 0, "an unreadable handler is not counted as framed");
    }

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
