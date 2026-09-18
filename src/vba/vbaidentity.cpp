#include "vbaidentity.h"
#include "core/text.h"

#include <windows.h>
#include <cstring>
#include <cstdio>
#include "core/safemem.h"

namespace vba
{
    namespace
    {
        // x64 layouts from the published prior art (Azzopardi's VBATrace), cross-checked by the
        // invariants each structure carries. Offsets, not addresses.

        // RTMI / p-code trailer
        constexpr std::uint32_t kRtmi_pParent      = 0x00;

        // parent
        constexpr std::uint32_t kPar_pObjTbl       = 0x08;
        constexpr std::uint32_t kPar_marker        = 0x20;   // == 0xFFFFFFFFFFFFFFFF
        constexpr std::uint32_t kPar_marker2       = 0x28;   // == 0
        constexpr std::uint32_t kPar_pListEntry    = 0x30;
        // WORD, not ULONG32. The published x64 struct declares this field and its
        // neighbour as 32-bit; on this build they are two 16-bit counts packed
        // together. Read from the live structure, not from the declaration.
        constexpr std::uint32_t kPar_nProcs        = 0x40;   // WORD
        constexpr std::uint32_t kPar_procMap       = 0x48;

        // module entry
        constexpr std::uint32_t kCe_pParent        = 0x00;   // back-pointer
        constexpr std::uint32_t kCe_marker         = 0x08;   // == 0xFFFFFFFFFFFFFFFF
        constexpr std::uint32_t kCe_pszModName     = 0x30;
        constexpr std::uint32_t kCe_nNumProcs      = 0x38;   // DWORD
        constexpr std::uint32_t kCe_ppszFnNames    = 0x40;

        // object table
        constexpr std::uint32_t kOt_pOwner         = 0x08;

        // owner (workbook)
        constexpr std::uint32_t kOwn_wszFilename   = 0x26;   // inline wchar_t[]

        constexpr std::uint64_t kMarker            = 0xFFFFFFFFFFFFFFFFull;
        constexpr std::uint32_t kMaxProcs          = 65536;  // sanity, not preference

        volatile LONG64 g_declines[static_cast<int>(IdDecline::Count_)] = {};

        // First failure only: the raw bytes, so a wrong offset is read rather than guessed.
        char           g_debug[1400] = {};
        volatile LONG  g_debugTaken = 0;

        void Decline(IdDecline d)
        {
            InterlockedIncrement64(&g_declines[static_cast<int>(d)]);
        }

        bool RdBlock(std::uint64_t at, std::uint64_t* dst, int qwords)
        {
            __try
            {
                const std::uint64_t* s = reinterpret_cast<const std::uint64_t*>(at);
                for (int i = 0; i < qwords; ++i) dst[i] = s[i];
                return true;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }
        }

        void CaptureDebug(std::uint64_t trailer, std::uint64_t parent, std::uint64_t listEntry)
        {
            if (InterlockedCompareExchange(&g_debugTaken, 1, 0) != 0) return;
            std::uint64_t q[16] = {};
            int n = 0;
            n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE,
                             "trailer=0x%llX parent=0x%llX listEntry=0x%llX\n",
                             (unsigned long long)trailer, (unsigned long long)parent,
                             (unsigned long long)listEntry);
            if (RdBlock(parent, q, 14))
            {
                n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE, "  parent   :");
                for (int i = 0; i < 14; ++i)
                    n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE,
                                     " +%02X=%llX", i * 8, (unsigned long long)q[i]);
                n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE, "\n");
            }
            if (RdBlock(listEntry, q, 12))
            {
                n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE, "  listEntry:");
                for (int i = 0; i < 12; ++i)
                    n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE,
                                     " +%02X=%llX", i * 8, (unsigned long long)q[i]);
                n += _snprintf_s(g_debug + n, sizeof(g_debug) - n, _TRUNCATE, "\n");
            }
        }

        // Address tests and guarded readers: core/safemem.h. A trailer is
        // 4-aligned, and the alignment is stated at each call site.
        using core::InRangeAndAligned;
        using core::RdU64;
        using core::RdU32;
        using core::RdU16;

        // A NUL-terminated narrow string, copied raw. `cut` says it did not fit.
        bool RdAnsiRaw(std::uint64_t at, char* dst, int cap, bool& cut)
        {
            cut = false;
            dst[0] = 0;
            __try
            {
                const char* s = reinterpret_cast<const char*>(at);
                int i = 0;
                for (; i < cap - 1 && s[i]; ++i) dst[i] = s[i];
                dst[i] = 0;
                cut = (i == cap - 1 && s[i]);
                return i > 0;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { dst[0] = 0; return false; }
        }

        // VBA keeps identifiers in the ANSI code page and the trace is UTF-8.
        // Truncation is marked with '~', never cut mid-character.
        bool RdAnsi(std::uint64_t at, char* dst, int cap)
        {
            if (cap < 2) { if (cap > 0) dst[0] = 0; return false; }
            dst[0] = 0;
            char raw[512];
            bool cut = false;
            if (!RdAnsiRaw(at, raw, static_cast<int>(sizeof raw), cut)) return false;

            char utf8[1536];
            int len = 0;
            bool ascii = true;
            for (const char* c = raw; *c; ++c)
                if (static_cast<unsigned char>(*c) >= 0x80) { ascii = false; break; }
            if (ascii)
            {
                len = static_cast<int>(strlen(raw));
                memcpy(utf8, raw, static_cast<size_t>(len));
            }
            else
            {
                wchar_t wide[512];
                const int wn = MultiByteToWideChar(CP_ACP, 0, raw, -1, wide, 512);
                len = (wn > 1) ? WideCharToMultiByte(CP_UTF8, 0, wide, wn - 1, utf8,
                                                     static_cast<int>(sizeof utf8), nullptr, nullptr)
                               : 0;
                if (len <= 0) return false;
            }

            const bool marked = cut || len > cap - 1;
            int keep = marked ? (len < cap - 2 ? len : cap - 2) : len;
            while (keep > 0 && keep < len && (static_cast<unsigned char>(utf8[keep]) & 0xC0) == 0x80) --keep;
            memcpy(dst, utf8, static_cast<size_t>(keep));
            if (marked) dst[keep++] = '~';
            dst[keep] = 0;
            return keep > 0;
        }

        bool RdWideAsAnsi(std::uint64_t at, char* dst, int cap)
        {
            if (cap <= 0) return false;
            dst[0] = 0;
            wchar_t tmp[260] = {};
            __try
            {
                const wchar_t* s = reinterpret_cast<const wchar_t*>(at);
                int i = 0;
                for (; i < 259 && s[i]; ++i) tmp[i] = s[i];
                tmp[i] = 0;
                if (i == 0) return false;
            }
            __except (EXCEPTION_EXECUTE_HANDLER) { return false; }

            // Leaf only: the full path is noise in a trace row, and the
            // workbook name is what a user recognises.
            const wchar_t* leaf = tmp;
            for (const wchar_t* p = tmp; *p; ++p)
                if (*p == L'\\' || *p == L'/') leaf = p + 1;

            return core::NarrowUtf8(leaf, -1, dst, cap) > 0;
        }
    }

    const char* IdDeclineName(IdDecline d)
    {
        switch (d)
        {
        case IdDecline::NullTrailer:         return "trailer was null";
        case IdDecline::TrailerUnreadable:   return "trailer unreadable";
        case IdDecline::ParentUnreadable:    return "parent unreadable";
        case IdDecline::ParentMarkerBad:     return "parent markers wrong";
        case IdDecline::ObjTableUnreadable:  return "object table unreadable";
        case IdDecline::ListEntryUnreadable: return "module entry unreadable";
        case IdDecline::BackPointerMismatch: return "module entry does not point back at the parent";
        case IdDecline::ProcCountInsane:     return "procedure count out of range";
        case IdDecline::ProcMapMiss:         return "trailer not found in the parent's procMap";
        case IdDecline::NameUnreadable:      return "name string unreadable";
        default:                             return "?";
        }
    }

    std::uint64_t IdDeclineCount(IdDecline d)
    {
        return static_cast<std::uint64_t>(g_declines[static_cast<int>(d)]);
    }

    void ResetIdentityCounts()
    {
        for (int i = 0; i < static_cast<int>(IdDecline::Count_); ++i) g_declines[i] = 0;
        g_debug[0] = 0;
        g_debugTaken = 0;
    }

    const char* IdentityDebug() { return g_debug; }

    // The same walk as Resolve, minus the requirement that a NAME come out of
    // it. Diagnostics need the addresses even when the name step fails, and a
    // walk that gives up early would hide exactly the structure being looked
    // for. Nothing here is on the hot path.

    bool Resolve(std::uint64_t trailer, Identity& out)
    {
        out = Identity{};
        if (!trailer) { Decline(IdDecline::NullTrailer); return false; }
        if (!InRangeAndAligned(trailer, 4)) { Decline(IdDecline::TrailerUnreadable); return false; }

        std::uint64_t parent = 0;
        if (!RdU64(trailer + kRtmi_pParent, parent) || !InRangeAndAligned(parent, 8))
        { Decline(IdDecline::ParentUnreadable); return false; }

        // The parent's two constant markers. This is the first real proof that
        // we are looking at the structure we think we are, rather than at
        // whatever happened to be at that address.
        std::uint64_t m1 = 0, m2 = 1;
        if (!RdU64(parent + kPar_marker, m1) || !RdU64(parent + kPar_marker2, m2))
        { Decline(IdDecline::ParentUnreadable); return false; }
        if (m1 != kMarker || m2 != 0) { Decline(IdDecline::ParentMarkerBad); return false; }

        std::uint64_t listEntry = 0, objTbl = 0, procMap = 0;
        std::uint16_t nProcs = 0;
        if (!RdU64(parent + kPar_pListEntry, listEntry) ||
            !RdU64(parent + kPar_pObjTbl,    objTbl)    ||
            !RdU64(parent + kPar_procMap,    procMap)   ||
            !RdU16(parent + kPar_nProcs,     nProcs))
        { Decline(IdDecline::ParentUnreadable); return false; }

        if (!InRangeAndAligned(listEntry, 8)) { Decline(IdDecline::ListEntryUnreadable); return false; }

        // The module entry's own marker, and -- the strongest check available --
        // its back-pointer to the parent we came from. Two independent
        // structures agreeing about each other is hard to fake by accident.
        std::uint64_t ceMarker = 0, ceParent = 0;
        if (!RdU64(listEntry + kCe_marker, ceMarker) ||
            !RdU64(listEntry + kCe_pParent, ceParent))
        { Decline(IdDecline::ListEntryUnreadable); return false; }
        if (ceMarker != kMarker) { Decline(IdDecline::ListEntryUnreadable); return false; }
        if (ceParent != parent)  { Decline(IdDecline::BackPointerMismatch); return false; }

        std::uint32_t nNames = 0;
        std::uint64_t pszMod = 0, ppszNames = 0;
        if (!RdU32(listEntry + kCe_nNumProcs,   nNames) ||
            !RdU64(listEntry + kCe_pszModName,  pszMod) ||
            !RdU64(listEntry + kCe_ppszFnNames, ppszNames))
        { Decline(IdDecline::ListEntryUnreadable); return false; }

        // Bound each array by its own length: nProcs for procMap, nNames for the names.
        const std::uint32_t mapCount  = static_cast<std::uint32_t>(nProcs);
        const std::uint32_t nameCount = nNames;
        if (mapCount == 0 || mapCount > kMaxProcs || nameCount > kMaxProcs)
        {
            CaptureDebug(trailer, parent, listEntry);
            Decline(IdDecline::ProcCountInsane);
            return false;
        }

        // Which procedure of this module are we? The parent's procMap is an
        // array of trailers, one per procedure, and our index into it is the
        // index of our name.
        int found = -1;
        if (InRangeAndAligned(procMap, 8))
        {
            for (std::uint32_t i = 0; i < mapCount; ++i)
            {
                std::uint64_t entry = 0;
                if (!RdU64(procMap + i * 8ull, entry)) break;
                if (entry == trailer) { found = static_cast<int>(i); break; }
            }
        }
        if (found < 0)
        {
            CaptureDebug(trailer, parent, listEntry);
            Decline(IdDecline::ProcMapMiss);
            return false;
        }

        // Names. A failure past this point still yields the fields we DID get.
        bool any = false;
        if (InRangeAndAligned(ppszNames, 8) && static_cast<std::uint32_t>(found) < nameCount)
        {
            std::uint64_t pName = 0;
            if (RdU64(ppszNames + static_cast<std::uint64_t>(found) * 8ull, pName) &&
                InRangeAndAligned(pName, 8))
                any = RdAnsi(pName, out.function, sizeof(out.function));
        }
        if (!any) { Decline(IdDecline::NameUnreadable); return false; }

        if (InRangeAndAligned(pszMod, 8)) RdAnsi(pszMod, out.module, sizeof(out.module));

        // Counted on both failure paths, so an unnamed workbook can be told from
        // one the walk failed to read.
        if (!InRangeAndAligned(objTbl, 8))
        {
            Decline(IdDecline::ObjTableUnreadable);
        }
        else
        {
            std::uint64_t owner = 0;
            if (RdU64(objTbl + kOt_pOwner, owner) && InRangeAndAligned(owner, 8))
                RdWideAsAnsi(owner + kOwn_wszFilename, out.workbook, sizeof(out.workbook));
            else
                Decline(IdDecline::ObjTableUnreadable);   // the owner half failed
        }

        out.ok = true;
        return true;
    }
}
