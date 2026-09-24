// Unit test for the last-statement check in vba::ReadArgTypes (vba/vbapcode.cpp): the last
// statement has no successor to land on, so its exit must end at ProcSize, after 0 or 2 bytes
// of padding. A walk that does not is not clean, and the disarm warning names the exit.
//
// One procedure in memory: a statement with no successor, then exit 635, then the trailer.
//
// Built by XRayXL.sln into build\x64\Release\unit\; the log and the corpus's module walk are
// stubbed below.

#include "vbapcode.h"
#include "vbatrailer.h"

#include <cstdio>
#include <cstring>
#include <string>

namespace core { namespace Log { void Warning(const std::string&) {} } }
namespace vba { int ModuleTrailers(std::uint64_t, std::uint64_t*, int) { return 0; } }   // the corpus is not tested here

namespace
{
    int g_fail = 0;
    void Check(bool ok, const char* what) { printf("[%s] %s\n", ok ? "PASS" : "FAIL", what); if (!ok) ++g_fail; }

    // The code, then the trailer ProcSize bytes after its start.
    alignas(8) std::uint8_t g_mem[512];

    bool Walk(std::uint16_t exitOp, std::uint16_t procSize, vba::ArgTypes& out)
    {
        std::memset(g_mem, 0, sizeof g_mem);
        const std::uint8_t code[] = { 0x67, 0x02, 0, 0, 0, 0,                  // 615, the last statement
                                      static_cast<std::uint8_t>(exitOp), static_cast<std::uint8_t>(exitOp >> 8) };
        std::memcpy(g_mem + 64, code, sizeof code);
        std::uint8_t* trailer = g_mem + 64 + procSize;
        std::memcpy(trailer + vba::kTrl_procSize, &procSize, 2);
        vba::ResetPcodeCounts();
        return vba::ReadArgTypes(reinterpret_cast<std::uint64_t>(trailer), out);
    }
}

int main()
{
    static vba::PcodeLengths L;
    L.slots = 1700; L.len[615] = 6; L.pinned = 1; L.ok = true;
    vba::SetArmedLengths(L);
    vba::ArgTypes out;

    Check(Walk(635, 8, out) && !out.partial, "an exit ending at ProcSize closes");
    Check(vba::PcodeClosureWarning()[0] == 0, "...and warns of nothing");

    Check(Walk(635, 10, out) && !out.partial, "an exit followed by 2 bytes of padding closes");

    const bool walked = Walk(635, 12, out);
    Check(walked && out.partial, "an exit 4 bytes short of ProcSize is not clean");
    const std::string w = vba::PcodeClosureWarning();
    printf("  %s\n", w.c_str());
    Check(w.find("op635=1") != std::string::npos, "...and the warning names the exit");

    Walk(953, 8, out);
    const std::string line = vba::PcodeLine();
    printf("  %s\n", line.c_str());
    Check(vba::PcodeClosureWarning()[0] == 0 && line.find("unmeasured length") != std::string::npos,
          "an exit of unmeasured length is counted as unchecked, not as closed or open");

    printf("\n%s\n", g_fail == 0 ? "ALL PASS" : "FAILED");
    return g_fail == 0 ? 0 : 1;
}
