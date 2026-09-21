// Turning a p-code trailer into a NAME.
//
// The trailer identifies a procedure but says nothing about what it is called.
// The name is reached by walking a chain of interpreter structures:
//
//   trailer(RTMI) -> parent -> objTable -> owner        workbook file name
//                           -> listEntry -> pszModName  module name
//                           -> procMap[j] == trailer    j = this procedure
//                           -> listEntry -> ppszFnNames[j]   the name
//
// Every step is unverified memory, so each is guarded and validated against the structures' own
// invariants: the parent and module entry carry the same marker, in one of two known forms; in
// the usual form the module entry points back at the parent; the two agree on the procedure
// count; and the trailer is in the parent's procMap. A failed check yields no name rather than
// a wrong one.
//
// Resolved once per procedure, when its trailer is first seen, and cached.
#pragma once
#include <cstdint>

namespace vba
{
    // Sized to VBA's own 255-character identifier limit: a truncated name is one a user cannot
    // search for. A temporary, filled by Resolve and copied into the procedure table.
    struct Identity
    {
        bool ok = false;
        // A file leaf name. Windows allows 255 characters.
        char workbook[256] = {};   // "Book1.xlsm"  (leaf, not the full path)
        // VBA's limit is 31, so 256 is not needed -- but this is a stack
        // temporary and uniformity is worth more than the bytes. A longer name
        // is reported by the '~' marker rather than decided by a buffer.
        char module[256]   = {};   // "Module1"
        // VBA's identifier limit: 255 characters, plus the terminator.
        char function[256] = {};   // "MyProcedure"
    };

    // Counted, never merged: "we never looked" and "we looked and the structure
    // did not check out" are different facts.
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

    // A one-shot hex dump of the first chain that failed to check out, so a
    // wrong offset can be READ rather than guessed at.
    const char* IdentityDebug();
}
