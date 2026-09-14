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
// EVERY STEP IS UNVERIFIED MEMORY UNTIL IT CHECKS OUT, so each is guarded and
// validated before the next is taken -- against the structures' own invariants:
// two constant markers, a reference count, and a back-pointer from the module
// entry to the parent. A chain failing any check yields NO NAME rather than a
// wrong one, because a confident wrong name is worse than an address.
//
// OFF THE HOT PATH: resolved once per procedure, when its trailer is first
// seen, and cached against that trailer.
#pragma once
#include <cstdint>

namespace vba
{
    // SIZED TO WHAT VBA ITSELF ALLOWS, not to what looked generous: VBA permits
    // a 255-character identifier, and a truncated name is one a user cannot
    // search their own code for.
    //
    // A TEMPORARY, filled by Resolve and copied into the procedure table, so
    // enlarging it costs stack rather than a table entry. The table's own
    // buffers are in vbatrace.cpp and are sized to match.
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
        ProcMapMiss,          // our trailer is not in the parent's procMap
        NameUnreadable,
        Count_
    };

    const char*   IdDeclineName(IdDecline d);
    std::uint64_t IdDeclineCount(IdDecline d);
    void          ResetIdentityCounts();

    // Resolve, or return false and leave `out` empty. Never throws.
    bool Resolve(std::uint64_t trailer, Identity& out);

    // A one-shot hex dump of the first chain that failed to check out, so a
    // wrong offset can be READ rather than guessed at.
    const char* IdentityDebug();
}
