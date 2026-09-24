// Turning a p-code trailer into a name, by walking the interpreter's structures:
//
//   trailer(RTMI) -> parent -> objTable -> owner        workbook file name
//                           -> listEntry -> pszModName  module name
//                           -> procMap[j] == trailer    j = this procedure
//                           -> listEntry -> ppszFnNames[j]   the name
//
// Every step is unverified memory, so each is guarded and checked against the structures' own
// invariants (marker, back-pointer, procedure count, procMap membership). A failed check yields
// no name rather than a wrong one.
#pragma once
#include <cstdint>

namespace vba
{
    // Sized to VBA's 255-character identifier limit: a truncated name is one a user cannot
    // search for.
    struct Identity
    {
        bool ok = false;
        char workbook[256] = {};   // "Book1.xlsm"  (leaf, not the full path)
        // VBA allows 31, but a longer name is marked '~' rather than decided by a buffer.
        char module[256]   = {};   // "Module1"
        char function[256] = {};   // "MyProcedure"
    };

    // Counted, never merged: "never looked" and "looked and it did not check out" differ.
    enum class IdDecline
    {
        NullTrailer,
        TrailerUnreadable,
        ParentUnreadable,
        ParentMarkerBad,      // marker/marker2 not the documented constants
        ObjTableUnreadable,
        ListEntryUnreadable,
        BackPointerMismatch,  // module entry does not point back at the parent
        ProcCountInsane,
        ProcCountMismatch,    // the parent's count is not the module entry's
        ProcMapMiss,          // our trailer is not in the parent's procMap
        NameUnreadable,
        Count_
    };

    const char*   IdDeclineName(IdDecline d);
    std::uint64_t IdDeclineCount(IdDecline d);
    void          ResetIdentityCounts();

    // Resolve, or return false and leave `out` empty. Never throws.
    bool Resolve(std::uint64_t trailer, Identity& out);

    // Every trailer in the module that owns `trailer`, run or not, after the same checks as
    // Resolve; null entries are kept. Counts no declines. Returns how many were written.
    int ModuleTrailers(std::uint64_t trailer, std::uint64_t* out, int cap);

    // A one-shot hex dump of the first chain that failed, so a wrong offset can be read.

    const char* IdentityDebug();
}
