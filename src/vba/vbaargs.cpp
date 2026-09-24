#include "vbaargs.h"
#include "vbaoleaut.h"
#include "vbapcode.h"
#include "vbaretdecode.h"
#include "vbatrailer.h"
#include <windows.h>
#include <cstdio>
#include <cstring>
#include <cstdlib>
#include "core/safemem.h"
#include "core/hash.h"
#include "core/valueformat.h"

namespace vba
{
    namespace
    {
        using core::RdU64;
        using core::RdU32;
        using core::RdU16;

        // Slots, not parameters: VBA's 60 parameters can need 181 when each is a `ByVal
        // Variant`. ArgTypes::kMax moves with it.
        constexpr int kMaxSlots = 192;

        volatile LONG64 g_declines[static_cast<int>(ArgDecline::Count_)] = {};
        volatile LONG64 g_captured = 0;
        // Not a decline, since what was recovered is correct, but still reported.
        volatile LONG64 g_partialWalks = 0;

        // `?opNNN` slots, which are holes in kTypeOps.
        constexpr int   kUnkMax = 16;
        volatile LONG   g_unkOp[kUnkMax]    = {};
        volatile LONG   g_unkCount[kUnkMax] = {};

        void NoteUnnamedArgOp(std::uint16_t op)
        {
            for (int i = 0; i < kUnkMax; ++i)
            {
                const LONG o = InterlockedCompareExchange(&g_unkOp[i], 0, 0);
                if (o == op) { InterlockedIncrement(&g_unkCount[i]); return; }
                if (o == 0 && InterlockedCompareExchange(&g_unkOp[i], op, 0) == 0)
                { InterlockedIncrement(&g_unkCount[i]); return; }
            }
        }

        void NoteDecline(ArgDecline d)
        {
            InterlockedIncrement64(&g_declines[static_cast<int>(d)]);
        }

        // Does not mask `vt & kVT_TYPEMASK`: every caller passes a vartype already resolved.
        using vba::VtName;


        // "Variant" is absent on purpose: the Variant decoder reads the tag itself. Compared over
        // `len`, so "Long&" is looked up without its "&".
        std::uint16_t VartypeFor(const char* name, size_t len)
        {
            struct Vt { const char* name; std::uint16_t vt; };
            static constexpr Vt kVt[] = {
                { "Byte",     17 },   // VT_UI1
                { "Integer",   2 },   // VT_I2 -- Boolean too
                { "Long",      3 },   // VT_I4
                { "Single",    4 },   // VT_R4
                { "Double",    5 },   // VT_R8 -- Date too
                { "Currency",  6 },   // VT_CY
                { "String",    8 },   // VT_BSTR
                { "Object",    9 },   // VT_DISPATCH
                { "LongLong", 20 },   // VT_I8
            };
            if (!name) return 0;
            for (const Vt& t : kVt)
                if (strlen(t.name) == len && strncmp(t.name, name, len) == 0) return t.vt;
            return 0;
        }

        // An omitted Optional is a Variant of VT_ERROR and DISP_E_PARAMNOTFOUND. Two exact
        // constants, so it is safe to decode with no type information.
        bool ReadMissingMarker(std::uint64_t p, core::ValueWriter& w)
        {
            if (p < 0x10000 || (p & 7)) return false;
            std::uint16_t vt = 0;
            std::uint32_t scode = 0;
            if (!RdU16(p, vt) || vt != kVT_ERROR) return false;
            if (!RdU32(p + 8, scode) || scode != kParamNotFound) return false;
            w.Word("Missing");
            return true;
        }

        // Marker first, so an omitted Optional reads `Missing`, not `Error(0x80020004)`; an
        // application Error() value carries a different scode.
        bool DescribeVariantOrMissing(std::uint64_t at, core::ValueWriter& w)
        {
            return ReadMissingMarker(at, w) || DescribeVariantValue(at, w);
        }

        // One indirection down for a `ByRef s As String`. Safe untyped because a BSTR proves
        // itself; `followed` makes the slot ByRef in substance.
        bool ReadBstrDirectOrThrough(std::uint64_t p, core::ValueWriter& w, bool& followed)
        {
            followed = false;
            if (DescribeBstrValue(p, w)) return true;
            // Range check first: a guarded read of a small integer costs an access violation.
            if (!core::InUserRange(p)) return false;
            std::uint64_t inner = 0;
            if (!RdU64(p, inner)) return false;
            if (!DescribeBstrValue(inner, w)) return false;
            followed = true;
            return true;
        }

        // VBA's descriptor for an empty ParamArray names no element type, so it is accepted only
        // when every field matches exactly, and reads `?[0..-1]{}`.
        bool ReadBlankEmptyArray(std::uint64_t psa, core::ValueWriter& w)
        {
            if (!core::InRangeAndAligned(psa, 8)) return false;
            std::uint64_t q[4] = {};
            for (int i = 0; i < 4; ++i)
                if (!RdU64(psa + static_cast<std::uint64_t>(i) * 8, q[i])) return false;
            // cDims=1, fFeatures=0, cbElements=0 | cLocks=0 | pvData=0 | cElements=0, lLbound=0
            if (q[0] != 1 || q[1] != 0 || q[2] != 0 || q[3] != 0) return false;
            const long long lo = 0, hi = -1;
            w.BeginArray(nullptr, 1, &lo, &hi);
            w.BeginLevel();
            w.EndLevel();
            w.EndArray();
            return true;
        }

        bool ReadSafeArrayValue(std::uint64_t p, core::ValueWriter& w)
        {
            if (ReadBlankEmptyArray(p, w)) return true;
            SaInfo s{};
            // One indirection down for `ByRef a()`. No vtHint: the descriptor must name its own
            // element type.
            if (!ReadSafeArrayHeader(p, 0, s))
            {
                if (!core::InUserRange(p)) return false;      // as above
                std::uint64_t inner = 0;
                if (!RdU64(p, inner) || !ReadSafeArrayHeader(inner, 0, s)) return false;
            }

            return RenderSafeArrayValue(s, w);
        }

        // One rule for the signature and the renderer, so they agree which slot is which argument.
        int SlotWidth(const char* tn)
        {
            return (tn && strcmp(tn, "Variant") == 0) ? 3 : 1;
        }

        // `tn` is the p-code's type, or null. False only when the slot cannot be read. Order:
        //   1. ByVal Variant: slots k.. are a VARIANT, bounded by argSz or the look-ahead
        //      reads the caller's frame.
        //   2. ByRef Variant: a VARIANT at the pointer.
        //   3. Declared scalar: never falls through, or a Double whose bits pass as a BSTR
        //      would render as text.
        //   4. UDT: its address. Before 5, since a String first field is a valid BSTR.
        //   5. Self-proving values: an omitted Optional, a BSTR, a SAFEARRAY.
        //   6. `Ref&` (747): arrays proved themselves in 5, so a 64-bit integer, except zero,
        //      which an unallocated array also reads as.
        //   7. The raw qword.
        bool RenderSlot(std::uint64_t r14, int k, int slots, const char* tn,
                        core::ValueWriter& w, bool& viaPointer)
        {
            const std::uint64_t slotAt = r14 + static_cast<std::uint64_t>(k) * 8;
            std::uint64_t v = 0;
            if (!RdU64(slotAt, v)) return false;

            const size_t tnLen   = tn ? strlen(tn) : 0;
            const bool   ref     = tnLen > 1 && tn[tnLen - 1] == '&';
            const size_t nameLen = ref ? tnLen - 1 : tnLen;
            auto declared = [&](const char* name)
            { return tn && strlen(name) == nameLen && strncmp(tn, name, nameLen) == 0; };

            // A reference that cannot be followed keeps no declared type.
            std::uint64_t shown = v, valueAt = slotAt;
            bool typed = (tn != nullptr);
            if (ref)
            {
                std::uint64_t inner = 0;
                if (RdU64(v, inner)) { shown = inner; valueAt = v; }
                else                 typed = false;
            }

            bool done = false, declaredScalar = false;
            if (!ref && declared("Variant"))                                    // 1
                done = (k + 1 < slots) && DescribeVariantOrMissing(slotAt, w);
            else if (typed && declared("Variant"))                              // 2
                done = DescribeVariantOrMissing(valueAt, w);
            else if (const std::uint16_t vt = typed ? VartypeFor(tn, nameLen) : 0) // 3
            {
                declaredScalar = true;
                done = DescribeArrayElement(vt, valueAt, w);
            }
            if (!done && typed && declared("Udt"))                              // 4
            {
                w.Udt(valueAt);
                done = true;
            }
            if (!done && !declaredScalar)                                       // 5
            {
                bool followed = false;
                done = ReadMissingMarker(shown, w)
                    || ReadBstrDirectOrThrough(shown, w, followed)
                    || ReadSafeArrayValue(shown, w);
                if (done && followed) viaPointer = true;
            }
            if (!done && typed && declared("Ref") && shown != 0)                // 6
            {
                char t[32];
                _snprintf_s(t, _TRUNCATE, "%lld", static_cast<long long>(static_cast<std::int64_t>(shown)));
                w.Number("LongLong", t);
                done = true;
            }
            if (!done)                                                          // 7
            {
                char t[32];
                _snprintf_s(t, _TRUNCATE, "0x%llX", static_cast<unsigned long long>(shown));
                w.Marker(t);
            }
            return true;
        }
    }

    const char* ArgDeclineName(ArgDecline d)
    {
        switch (d)
        {
        case ArgDecline::NoTrailer:           return "no trailer";
        case ArgDecline::NoFrameBase:         return "no frame base (r14)";
        case ArgDecline::ArgSzUnreadable:     return "argSz unreadable";
        case ArgDecline::ArgSzNotMultipleOf8: return "argSz not a multiple of 8";
        case ArgDecline::ArgSzOutOfRange:     return "argSz out of range";
        case ArgDecline::SlotUnreadable:      return "an argument slot faulted";
        case ArgDecline::NoRenderBuffer:      return "no render buffer";
        case ArgDecline::TooLarge:            return "arguments over the value limit";
        default:                              return "?";
        }
    }


    void ResetArgCounts()
    {
        for (int i = 0; i < static_cast<int>(ArgDecline::Count_); ++i) g_declines[i] = 0;
        g_captured = 0;
        g_partialWalks = 0;
        for (int i = 0; i < kUnkMax; ++i)
        { InterlockedExchange(&g_unkOp[i], 0); InterlockedExchange(&g_unkCount[i], 0); }
    }

    bool g_diagOps = false;

    // One function, so the signature and the value cell agree.
    //
    //    `?opNNN`  an opcode touched the slot and is not in the type table
    //    `?none`   an opcode touched it and is known to convey no type
    //    `?unseen` no opcode touched this slot at all
    //
    // `count` only for the signature, so totals do not double-count.
    static const char* UnnamedTypeMarker(const ArgTypes& types, int k, char* buf, int cap,
                                  bool count)
    {
        if (!types.op[k]) { _snprintf_s(buf, cap, _TRUNCATE, "?unseen"); return buf; }
        if (!PcodeCarriesNoType(types.op[k]))
        {
            _snprintf_s(buf, cap, _TRUNCATE, "?op%u", static_cast<unsigned>(types.op[k]));
            if (count) NoteUnnamedArgOp(types.op[k]);
        }
        else if (g_diagOps)
            _snprintf_s(buf, cap, _TRUNCATE, "?none#%u", static_cast<unsigned>(types.op[k]));
        else _snprintf_s(buf, cap, _TRUNCATE, "?none");
        return buf;
    }

    // Also sets out.params: one entry per parameter, not per slot. Needs out.slots and firstSlot.
    static void BuildSignature(const ArgTypes& types, bool haveTypes, ArgCapture& out)
    {
        // An empty signature with argcount 0 is how "no parameters" reads on both sources.
        if (out.slots <= 0) { out.params = 0; out.signature[0] = 0; return; }

        // No types recovered, so parameters cannot be told from slots.
        if (!haveTypes) { out.params = out.slots; return; }

        constexpr int cap = static_cast<int>(sizeof(out.signature));
        int  sj = 0;
        out.signature[0] = 0;
        int  params = 0;
        bool truncated = false;
        for (int k = out.firstSlot; k < out.firstSlot + out.slots && k < ArgTypes::kMax;
             k += SlotWidth(types.name[k]))
        {
            const char* tn = types.name[k];
            char unk[24];
            if (!tn) tn = UnnamedTypeMarker(types, k, unk, sizeof unk, /*count=*/true);
            char named[40];
            if (g_diagOps && types.name[k])
            {
                _snprintf_s(named, _TRUNCATE, "%s#%u", tn, static_cast<unsigned>(types.op[k]));
                tn = named;
            }
            // Every parameter is counted, written or not: argcount is the count, not the text.
            ++params;
            if (truncated) continue;
            const int w = _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                                      "%s%s", params > 1 ? "," : "", tn);
            // End with ",..." rather than close short, which would read as fewer parameters.
            if (w < 0 || sj + w > cap - 6) { truncated = true; out.signature[sj] = 0; continue; }
            sj += w;
        }

        if (g_diagOps && types.exitOp && sj < cap - 16)
        {
            const int wx = _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                                       "|exit%u", static_cast<unsigned>(types.exitOp));
            if (wx > 0) sj += wx;
        }

        // `~`: a `?` may be a skipped load rather than an unread parameter.
        if (types.partial) InterlockedIncrement64(&g_partialWalks);
        _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                    truncated ? ",..." : (types.partial ? "~" : ""));
        out.params = params;
    }

    void SetArgTypeOpcodeDiagnostics(bool on) { g_diagOps = on; }

    bool CaptureArgs(std::uint64_t trailer, std::uint64_t r14, ArgCapture& out)
    {
        out.ok = false; out.slots = 0; out.byRefShapeOnly = false;
        if (!out.text)
        {
            NoteDecline(ArgDecline::NoRenderBuffer);
            return false;
        }
        core::TextBuf& buf = *out.text;
        buf.Clear();

        if (!trailer) { NoteDecline(ArgDecline::NoTrailer);   return false; }
        if (!r14)     { NoteDecline(ArgDecline::NoFrameBase); return false; }

        std::uint16_t argSz = 0;
        if (!RdU16(trailer + kTrl_argSz, argSz))
        {
            NoteDecline(ArgDecline::ArgSzUnreadable);
            return false;
        }
        if (argSz & 7)      { NoteDecline(ArgDecline::ArgSzNotMultipleOf8); return false; }
        // Slot 0 is reserved, so a valid region is at least one slot.
        const int slots = argSz / 8;
        if (slots < 1 || slots > kMaxSlots)
        {
            NoteDecline(ArgDecline::ArgSzOutOfRange);
            return false;
        }

        // Bounded before any result slot is taken off: an index past that is provably wrong.
        ArgTypes types;
        const bool haveTypes = ReadArgTypes(trailer, types, slots - 1);

        // A Function returning a Variant or a record gets the caller's result buffer first.
        const RetKind rk = ExitReturnKind(types.exitOp);
        out.firstSlot = (rk == RetKind::Variant || rk == RetKind::Record) ? 2 : 1;
        out.slots     = slots - out.firstSlot;

        // A class or form Function returns through a trailing [out, retval] slot. Taken off only
        // when the exit's operand names that slot.
        if (out.slots > 0 && ExitHasTrailingResultSlot(types.exitOp) &&
            types.exitOperand == (slots - 1) * 8)
        {
            --out.slots;
        }
        if (out.slots < 0) out.slots = 0;

        BuildSignature(types, haveTypes, out);

        // The label a<n> counts slots, so after a ByVal Variant at a1 the next parameter is a4.
        std::uint64_t refHash = core::kFnvOffset;
        core::WithValueWriter(buf, [&](core::ValueWriter& w)
        {
            w.BeginArgs();
            for (int k = out.firstSlot; k < out.firstSlot + out.slots && k < ArgTypes::kMax; )
            {
                const int   n  = k - out.firstSlot + 1;
                const char* tn = haveTypes ? types.name[k] : nullptr;
                const size_t tnLen = tn ? strlen(tn) : 0;
                const bool  declaredRef = tnLen > 1 && tn[tnLen - 1] == '&';
                // `tn` stays null into RenderSlot: the marker is a label, never a type to decode as.
                char unk[24];
                const char* label = tn ? tn : UnnamedTypeMarker(types, k, unk, sizeof unk,
                                                                /*count=*/false);
                // A parameter only passed on has no type of its own; a reader takes it from the
                // callee's row with the same address.
                const std::uint64_t slotAt = r14 + static_cast<std::uint64_t>(k) * 8;
                std::uint64_t address = 0;
                if (declaredRef || (!tn && PcodePassesHeldPointer(types.op[k])))
                    RdU64(slotAt, address);
                else if (!tn && PcodePassesSlotAddress(types.op[k]))
                    address = slotAt;
                const core::ValueWriter::Mark before = w.Save();
                const int refusedBefore = buf.refused;
                w.BeginArg(n, label, address);
                const std::size_t valueAt = buf.Len();
                bool followed = false;
                if (!RenderSlot(r14, k, slots, tn, w, followed))
                {
                    NoteDecline(ArgDecline::SlotUnreadable);
                    if (!out.byRefOnly || declaredRef) { w.ArgUnreadable(); w.EndArg(); }
                    else                               w.Restore(before);
                    break;
                }
                if (followed) out.viaPointer = true;
                const bool isRef = declaredRef || followed;
                if (out.byRefOnly && !isRef) { w.Restore(before); k += SlotWidth(tn); continue; }
                // Not the separators, which depend on what was written before.
                if (isRef)
                {
                    if (buf.refused > refusedBefore) out.byRefShapeOnly = true;
                    refHash = core::Fnv1aByte(refHash, 0x1F);
                    refHash = core::Fnv1aByte(refHash, static_cast<std::uint8_t>(n));
                    for (const char* c = label; *c; ++c)
                        refHash = core::Fnv1aByte(refHash, static_cast<std::uint8_t>(*c));
                    for (std::size_t q = valueAt; q < buf.Len(); ++q)
                        refHash = core::Fnv1aByte(refHash, static_cast<std::uint8_t>(buf.Text()[q]));
                }
                w.EndArg();
                k += SlotWidth(tn);
            }
            w.EndArgs();
        });
        // The exit re-read is compared, never written, so its refusals are not the file's.
        if (out.byRefOnly && buf.refused > 0) core::g_valuesRefused.fetch_sub(buf.refused);

        // Only an array refuses itself when it does not fit; anything else past the limit
        // leaves text that cannot be trusted, so none of it is kept.
        if (buf.Over())
        {
            buf.Clear();
            NoteDecline(ArgDecline::TooLarge);
            return false;
        }

        out.byRefHash = refHash;
        out.ok = true;
        InterlockedIncrement64(&g_captured);
        return true;
    }

    const char* ArgTypeUnknownWarning()
    {
        static char b[512];
        int j = 0, n = 0;
        for (int i = 0; i < kUnkMax; ++i)
        {
            const LONG op = InterlockedCompareExchange(&g_unkOp[i], 0, 0);
            if (!op) continue;
            if (!n)
                j = _snprintf_s(b, _TRUNCATE,
                                "VBA args: opcode(s) touched a PARAMETER SLOT and are "
                                "not in the type table, so the parameter's declared type "
                                "was rendered '?opNNN' rather than named --");
            const int w = _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, " op%ld=%ld",
                                      op, InterlockedCompareExchange(&g_unkCount[i], 0, 0));
            if (w < 0) break;
            j += w; ++n;
        }
        if (!n) { b[0] = 0; return b; }
        _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                    ". Each is one row in vbapcode.cpp kTypeOps, and its length has to be "
                    "settled before the type can be named. Please report these with "
                    "the log.");
        return b;
    }

    const char* ArgsLine()
    {
        static char b[400];
        int j = _snprintf_s(b, _TRUNCATE, "VBA args: %llu captured",
                            static_cast<unsigned long long>(g_captured));
        if (g_partialWalks)
            j += _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE,
                             ", WARNING %llu signature(s) INCOMPLETE -- the walk"
                             " stopped early or resynchronised past a statement"
                             " (marked \"~\" -- a '?' in one of those may be a"
                             " parameter we failed to reach, not one the body"
                             " never reads)",
                             static_cast<unsigned long long>(g_partialWalks));
        for (int i = 0; i < static_cast<int>(ArgDecline::Count_); ++i)
        {
            if (!g_declines[i]) continue;
            const int w = _snprintf_s(b + j, sizeof(b) - j, _TRUNCATE, ", %s=%llu",
                                      ArgDeclineName(static_cast<ArgDecline>(i)),
                                      static_cast<unsigned long long>(g_declines[i]));
            if (w < 0) break;
            j += w;
        }
        return b;
    }
}
