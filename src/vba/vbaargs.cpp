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

namespace vba
{
    namespace
    {
        // Guarded readers: core/safemem.h.
        using core::RdU64;
        using core::RdU32;
        using core::RdU16;
        // kTrl_argSz (trailer+0x08, a WORD in bytes) is in vbatrailer.h.

        // A procedure with more than this many argument SLOTS is not believed.
        // Slots, not parameters: a `ByVal Variant` spans three, so VBA's maximum of
        // 60 parameters can need 181. ArgTypes::kMax is indexed by slot and moves with it.
        constexpr int kMaxSlots = 192;

        // Per-thread heap: the hooks run where VBA has the least stack left.
        constexpr int kSlotRenderMax = 16384;
        __declspec(thread) char* t_slotText    = nullptr;
        __declspec(thread) char* t_variantText = nullptr;
        char* ThreadText(char*& buf)
        {
            if (!buf) buf = static_cast<char*>(malloc(kSlotRenderMax));
            return buf;
        }

        volatile LONG64 g_declines[static_cast<int>(ArgDecline::Count_)] = {};
        volatile LONG64 g_captured = 0;
        // Signatures cut short by a walk that stopped. Not a decline -- what
        // was recovered is correct -- but information we expected and did
        // not get, so it is reported rather than absorbed.
        volatile LONG64 g_partialWalks = 0;

        // Opcodes that touched a parameter slot and could not be named. A bare `?` means the
        // body never loads the parameter; `?opNNN` is a hole in kTypeOps, so a session counts
        // those and says so.
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


        // The element cap is vbaretdecode.h's, shared.
        using vba::kMaxRenderedElems;

        // ONE VARTYPE TABLE, in vbaoleaut.h. It does NOT mask `vt & kVT_TYPEMASK`; every
        // caller here passes a vartype EffectiveElemVt has already resolved.
        using vba::VtName;


        // The declared type name as a VARTYPE for the shared decoder. "Variant" is absent on
        // purpose: the Variant decoder reads the tag itself. Unmapped names return 0. Compared
        // over `len`, so "Long&" is looked up without its "&".
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

        // An omitted Optional is materialised by the caller as a Variant carrying VT_ERROR and
        // DISP_E_PARAMNOTFOUND. Both constants are exact, so this is safe to decode with no
        // type information.

        // AN OMITTED Optional, OR NOTHING. The strictest reader here: it
        // demands two EXACT constants -- vt == VT_ERROR and scode ==
        // DISP_E_PARAMNOTFOUND -- and nothing else reads as that pair.
        bool ReadMissingMarker(std::uint64_t p, char* out, int cap)
        {
            if (p < 0x10000 || (p & 7)) return false;
            std::uint16_t vt = 0;
            std::uint32_t scode = 0;
            if (!RdU16(p, vt) || vt != kVT_ERROR) return false;
            if (!RdU32(p + 8, scode) || scode != kParamNotFound) return false;
            _snprintf_s(out, cap, _TRUNCATE, "Missing");
            return true;
        }

        // An omitted Optional is itself a VARIANT, so the marker is asked first and reads
        // `Missing` rather than `Error(0x80020004)`. Safe because it demands two exact
        // constants: an application Error() value carries a different scode.
        bool DescribeVariantOrMissing(std::uint64_t at, char* out, int cap, int maxElems)
        {
            if (ReadMissingMarker(at, out, cap)) return true;
            const char* held = "";
            char* const vbuf = ThreadText(t_variantText);
            if (!vbuf || !DescribeVariantValue(at, vbuf, kSlotRenderMax, &held, maxElems))
                return false;
            // `args` has no type column beside it, so the held type is named
            // inline. Empty when the text already says it.
            if (held[0]) _snprintf_s(out, cap, _TRUNCATE, "%s(%s)", held, vbuf);
            else         _snprintf_s(out, cap, _TRUNCATE, "%s", vbuf);
            return true;
        }

        // A BSTR, direct or one indirection down: a `ByRef s As String` slot points to one.
        // Safe because a BSTR proves itself by length prefix, NUL termination and no control
        // characters. `followed` says a pointer was followed, which makes the slot ByRef in
        // substance.
        bool ReadBstrDirectOrThrough(std::uint64_t p, char* out, int cap, bool& followed)
        {
            followed = false;
            if (DescribeBstrValue(p, out, cap)) return true;
            // Range check first: a guarded read of a small integer costs an access violation.
            if (!core::InUserRange(p)) return false;
            std::uint64_t inner = 0;
            if (!RdU64(p, inner)) return false;
            if (!DescribeBstrValue(inner, out, cap)) return false;
            followed = true;
            return true;
        }

        // A SAFEARRAY AT `p` (or one indirection down -- a `ByRef a()` parameter),
        // OR NOTHING. The structure describes itself, and every part of that
        // description is checked before anything is read through it.
        bool ReadSafeArrayValue(std::uint64_t p, char* out, int cap)
        {
            SaInfo s{};
            // DIRECT, OR ONE INDIRECTION DOWN -- a `ByRef a()` parameter points at
            // the SAFEARRAY pointer. No vtHint: nothing on this path has named an
            // element type, so the descriptor must say so itself.
            if (!ReadSafeArrayHeader(p, 0, s))
            {
                if (!core::InUserRange(p)) return false;      // as above
                std::uint64_t inner = 0;
                if (!RdU64(p, inner) || !ReadSafeArrayHeader(inner, 0, s)) return false;
            }

            // One array renderer, shared with the return column (vbaretdecode.h).
            return RenderSafeArrayValue(s, out, cap, kMaxRenderedElems);
        }

        // Frame slots per parameter: one, or three for a ByVal Variant (a
        // 24-byte VARIANT). One rule for the signature and the value renderer,
        // so they cannot disagree about which slot is which argument.
        int SlotWidth(const char* tn)
        {
            return (tn && strcmp(tn, "Variant") == 0) ? 3 : 1;
        }

        // One argument slot as text. `tn` is the type the p-code named, or null. False only
        // when the slot cannot be read. Order matters:
        //   1. ByVal Variant: slots k and k+1 ARE a VARIANT, read by the return
        //      decoder so both columns render identical bytes identically.
        //      Bounded by argSz, or the look-ahead reads the caller's frame.
        //   2. ByRef Variant: a VARIANT at the pointer, read from its address.
        //   3. Declared scalar: the shared element decoder. A declared scalar
        //      never falls through to a structure reader -- a Double whose
        //      bits satisfy a BSTR's invariants must not render as text.
        //   4. A UDT is a record with no scalar value: its address. Before the
        //      probes, because a first field that is a String is a valid BSTR.
        //   5. Nothing declared, or a reference whose target proves itself:
        //      an omitted Optional, a BSTR, a SAFEARRAY. Following a pointer
        //      to find one marks the slot ByRef in substance, which the exit
        //      re-read needs for a write-only String.
        //   6. `Ref&` (747) is "eight bytes by reference" and nothing more; the
        //      array case proved itself in 4, what remains is a 64-bit integer.
        //   7. The raw qword.
        bool RenderSlot(std::uint64_t r14, int k, int slots, const char* tn,
                        char* one, int cap, bool& viaPointer)
        {
            const std::uint64_t slotAt = r14 + static_cast<std::uint64_t>(k) * 8;
            std::uint64_t v = 0;
            if (!RdU64(slotAt, v)) return false;

            // A name ending in "&" is a pointer to its target, and is compared
            // without the "&" from here on.
            const size_t tnLen   = tn ? strlen(tn) : 0;
            const bool   ref     = tnLen > 1 && tn[tnLen - 1] == '&';
            const size_t nameLen = ref ? tnLen - 1 : tnLen;
            auto declared = [&](const char* name)
            { return tn && strlen(name) == nameLen && strncmp(tn, name, nameLen) == 0; };

            // ByVal: the slot holds the value. ByRef: it holds a pointer, the
            // value lives there and `shown` is its first qword. A reference
            // that cannot be followed keeps no declared type.
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
                done = (k + 1 < slots) &&
                       DescribeVariantOrMissing(slotAt, one, cap, kMaxRenderedElems);
            else if (typed && declared("Variant"))                              // 2
                done = DescribeVariantOrMissing(valueAt, one, cap, kMaxRenderedElems);
            else if (const std::uint16_t vt = typed ? VartypeFor(tn, nameLen) : 0) // 3
            {
                declaredScalar = true;
                done = DescribeArrayElement(vt, valueAt, one, cap);
            }
            if (!done && typed && declared("Udt"))                              // 4
            {
                _snprintf_s(one, cap, _TRUNCATE, "udt@0x%llX",
                            static_cast<unsigned long long>(valueAt));
                done = true;
            }
            if (!done && !declaredScalar)                                       // 5
            {
                bool followed = false;
                done = ReadMissingMarker(shown, one, cap)
                    || ReadBstrDirectOrThrough(shown, one, cap, followed)
                    || ReadSafeArrayValue(shown, one, cap);
                if (done && followed) viaPointer = true;
            }
            if (!done && typed && declared("Ref"))                              // 6
            {
                _snprintf_s(one, cap, _TRUNCATE, "%lld",
                            static_cast<long long>(static_cast<std::int64_t>(shown)));
                done = true;
            }
            if (!done)                                                          // 7
                _snprintf_s(one, cap, _TRUNCATE, "0x%llX",
                            static_cast<unsigned long long>(shown));
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

    // XRAYXL_DIAG: show the opcode behind a name (see vbaargs.h).
    bool g_diagOps = false;

    // What to call a slot the bytecode did not name. One function, so the signature and the
    // value cell agree.
    //
    //    `?opNNN`  an opcode touched the slot and is not in the type table
    //    `?none`   an opcode touched it and is known to convey no type
    //    `?unseen` no opcode touched this slot at all
    //
    // `~` closing the signature means the walk did not read the whole body, so a `?unseen` may
    // be a load the walk skipped. `count` only where the marker reaches a signature, so the
    // totals do not double-count.
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

    // "(Type,Type,...)" from the recovered types, and out.params, the declared
    // parameter count: one entry per parameter, not per slot, stepping by the
    // same SlotWidth the value renderer uses. Needs out.slots and out.firstSlot.
    static void BuildSignature(const ArgTypes& types, bool haveTypes, ArgCapture& out)
    {
        // The slot COUNT comes from argSz arithmetic, not from the type walk,
        // so it is known even when no type is. Zero slots: an empty signature
        // with argcount 0, which is how "no parameters" reads on both sources.
        if (out.slots <= 0) { out.params = 0; out.signature[0] = 0; return; }

        // No types recovered, so parameters cannot be told from slots.
        if (!haveTypes) { out.params = out.slots; return; }

        constexpr int cap = static_cast<int>(sizeof(out.signature));
        int  sj = 0;
        out.signature[0] = 0;
        int  params = 0;
        bool truncated = false;
        // From the first real argument: slot 1 is the caller's result VARIANT
        // when firstSlot is 2, and not a parameter.
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
            const int w = _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                                      "%s%s", params ? "," : "", tn);
            // Out of buffer: say the signature was cut rather than close it
            // short, which would read as fewer parameters than there are.
            if (w < 0) { truncated = true; break; }
            sj += w;
            ++params;
        }
        // Keep room for the marker, or the cut is invisible again.
        if (truncated && sj > cap - 6) sj = cap - 6;

        // XRAYXL_DIAG: name the exit the walk ended on, so the opcode that
        // terminates a procedure can be read straight off a trace.
        if (g_diagOps && types.exitOp && sj < cap - 16)
        {
            const int wx = _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                                       "|exit%u", static_cast<unsigned>(types.exitOp));
            if (wx > 0) sj += wx;
        }

        // `~`: the walk did not read the whole body, so a `?` may be a
        // skipped load rather than an unread parameter.
        if (types.partial) InterlockedIncrement64(&g_partialWalks);
        _snprintf_s(out.signature + sj, cap - sj, _TRUNCATE,
                    truncated ? ",..." : (types.partial ? "~" : ""));
        out.params = params;
    }

    void SetArgTypeOpcodeDiagnostics(bool on) { g_diagOps = on; }

    bool CaptureArgs(std::uint64_t trailer, std::uint64_t r14, ArgCapture& out)
    {
        out.ok = false; out.slots = 0;
        // The caller owns the render buffer. No buffer -> decline cleanly:
        // a failed per-thread allocation must not fault the hook.
        char* const one = ThreadText(t_slotText);
        if (!out.text || out.textCap < 2 || !one)
        {
            NoteDecline(ArgDecline::NoRenderBuffer);
            return false;
        }
        out.text[0] = 0;

        if (!trailer) { NoteDecline(ArgDecline::NoTrailer);   return false; }
        if (!r14)     { NoteDecline(ArgDecline::NoFrameBase); return false; }

        std::uint16_t argSz = 0;
        if (!RdU16(trailer + kTrl_argSz, argSz))
        {
            NoteDecline(ArgDecline::ArgSzUnreadable);
            return false;
        }
        if (argSz & 7)      { NoteDecline(ArgDecline::ArgSzNotMultipleOf8); return false; }
        // Slot 0 is reserved, so a valid region is at least one slot and the
        // argument count is one less than the slot count.
        const int slots = argSz / 8;
        if (slots < 1 || slots > kMaxSlots)
        {
            NoteDecline(ArgDecline::ArgSzOutOfRange);
            return false;
        }

        // Types come from the bytecode, never from the bytes. The bound handed
        // to the walk is the slot count before any result slot is taken off:
        // a recovered index past it is provably wrong and is discarded.
        ArgTypes types;
        const bool haveTypes = ReadArgTypes(trailer, types, slots - 1);

        // How many slots are arguments: argSz = 8 x (nargs + 1) for the reserved slot 0, plus
        // one when a Function returns Variant and the caller's result VARIANT arrives first.
        // The return kind comes from the exit opcode the walk stopped on; with no typed exit
        // the plain arithmetic stands.
        out.firstSlot = (ExitReturnKind(types.exitOp) == RetKind::Variant) ? 2 : 1;
        out.slots     = slots - out.firstSlot;

        // A class or form Function is a COM method: its result comes back
        // through a trailing [out, retval] slot, so the last slot is not a
        // parameter. Checked against the exit instruction's own operand, the
        // byte offset of that slot; if they disagree nothing is taken away.
        if (out.slots > 0 && ExitHasTrailingResultSlot(types.exitOp) &&
            types.exitOperand == (slots - 1) * 8)
        {
            --out.slots;
        }
        if (out.slots < 0) out.slots = 0;

        BuildSignature(types, haveTypes, out);

        // `k` is the slot; the label a<n> counts slots from the first argument,
        // so a three-slot ByVal Variant at a1 puts the next parameter at a4 --
        // the label is a position. 16 KB so an array argument renders in full.
        std::uint64_t refHash = core::kFnvOffset;
        int j = 0;
        for (int k = out.firstSlot; k < out.firstSlot + out.slots && k < ArgTypes::kMax; )
        {
            const int   n  = k - out.firstSlot + 1;
            const char* tn = haveTypes ? types.name[k] : nullptr;
            const size_t tnLen = tn ? strlen(tn) : 0;
            const bool  declaredRef = tnLen > 1 && tn[tnLen - 1] == '&';
            // A named type is VBA's own metadata and the value was read as that type; a marker
            // means the value was recognised from the bytes themselves. `tn` stays null into
            // RenderSlot: the marker is a label, never a type to decode against.
            char unk[24];
            const char* label = tn ? tn : UnnamedTypeMarker(types, k, unk, sizeof unk,
                                                            /*count=*/false);
            bool followed = false;
            if (!RenderSlot(r14, k, slots, tn, one, kSlotRenderMax, followed))
            {
                // Keep what is known, mark what is not.
                NoteDecline(ArgDecline::SlotUnreadable);
                if (!out.byRefOnly || declaredRef)
                    _snprintf_s(out.text + j, out.textCap - j, _TRUNCATE, "%sa%d:%s=<unreadable>",
                                j ? " " : "", n, label);
                break;
            }
            if (followed) out.viaPointer = true;
            const bool isRef = declaredRef || followed;
            if (out.byRefOnly && !isRef) { k += SlotWidth(tn); continue; }
            const int lead = j ? 1 : 0;
            const int w = _snprintf_s(out.text + j, out.textCap - j, _TRUNCATE, "%sa%d:%s=%s",
                                      lead ? " " : "", n, label, one);
            if (w < 0) break;
            // Hashed without the separator, which depends on what was rendered before it.
            if (isRef)
            {
                refHash = core::Fnv1aByte(refHash, 0x1F);
                for (int q = lead; q < w; ++q)
                    refHash = core::Fnv1aByte(refHash, static_cast<std::uint8_t>(out.text[j + q]));
            }
            j += w;
            k += SlotWidth(tn);
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
        // [length first: candidate consensus over the shape corpus]
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
