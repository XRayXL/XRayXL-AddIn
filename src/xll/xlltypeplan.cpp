#include "xlltypeplan.h"

namespace xll
{
    namespace
    {
        // Map one type code (already assembled, so "K" and "K%" are distinct)
        // to what the ABI slot holds.
        Kind KindOf(wchar_t c, bool wide)
        {
            switch (c)
            {
            case L'A': return Kind::Int16;          // boolean as short
            case L'B': return Kind::Double;         // the only XMM case
            case L'C': return wide ? Kind::StrWideZ     : Kind::StrAsciiZ;
            case L'D': return wide ? Kind::StrWideCount : Kind::StrAsciiCount;
            case L'E': return Kind::PtrDouble;
            case L'F': return wide ? Kind::StrWideZ     : Kind::StrAsciiZ;      // modify-in-place
            case L'G': return wide ? Kind::StrWideCount : Kind::StrAsciiCount;  // modify-in-place
            case L'H': return Kind::UInt16;         // unsigned short
            case L'I': return Kind::Int16;          // signed short
            case L'J': return Kind::Int32;
            case L'K': return wide ? Kind::Fp12 : Kind::Fp;
            case L'L': return Kind::PtrInt16;       // pointer to boolean
            case L'M': return Kind::PtrInt16;
            case L'N': return Kind::PtrInt32;
            case L'O': return Kind::ArrayTriple;    // THREE slots
            case L'P': return Kind::OperNarrow;
            case L'Q': return Kind::OperWide;
            case L'R': return Kind::OperNarrow;
            case L'U': return Kind::OperWide;
            case L'X': return Kind::AsyncHandle;
            default:   return Kind::Unknown;
            }
        }

        // Trailing flags: $ thread-safe, ! volatile, # macro-sheet, & cluster-safe.
        bool IsModifier(wchar_t c)
        {
            return c == L'!' || c == L'#' || c == L'$' || c == L'&';
        }

        // No register return: '>' is void, digits modify in place, F/G return via their buffer.
        bool ReturnIsVoid(wchar_t c)
        {
            return c == L'>' || c == L'F' || c == L'G' || (c >= L'1' && c <= L'9');
        }

        // The C API gives only these codes a wide form, as argument or return.
        bool TakesWideSuffix(wchar_t c)
        {
            return c == L'C' || c == L'D' || c == L'F' || c == L'G' || c == L'K' || c == L'O';
        }
    }

    Plan Parse(const wchar_t* typeText)
    {
        Plan p;
        if (typeText == nullptr || typeText[0] == 0) return p;

        int i = 0;
        char flags[5] = {};
        int  nflags = 0;
        int  sigLen = 0;
        bool sigCut = false;
        auto sigPut = [&](const char* s)
        {
            for (; *s && sigLen < static_cast<int>(sizeof(p.signature)) - 1; ++s) p.signature[sigLen++] = *s;
            p.signature[sigLen] = 0;
        };

        // ---- the return code -------------------------------------------------
        // '>' is the return code itself, not a prefix: ">QX" is void f(Q, X).
        {
            const wchar_t c = typeText[i];
            const bool wide = (typeText[i + 1] == L'%');
            if (wide && !TakesWideSuffix(c)) { p.firstBadCode = '%'; return p; }

            p.returnCode[0] = static_cast<char>(c);
            if (wide) p.returnCode[1] = '%';

            if (ReturnIsVoid(c))
            {
                // Nothing to read out of a register. returnKind == Void is the
                // signal DescribeReturn reads.
                p.returnKind = Kind::Void;
            }
            else
            {
                p.returnKind = KindOf(c, wide);
                if (p.returnKind == Kind::Unknown)
                {
                    p.firstBadCode = static_cast<char>(c);
                    return p;   // ok stays false
                }
            }
            i += wide ? 2 : 1;
        }

        // `typeIndex` walks the string and `abi` walks the ABI slots, and they differ: a '%' is
        // part of its code, and an O takes three slots.
        int abi = 0;
        for (; typeText[i] != 0; )
        {
            const wchar_t c = typeText[i];

            if (IsModifier(c))
            {
                // Registration flags, not arguments: the slot walk skips them and rettype carries them.
                if (nflags < 4) flags[nflags++] = static_cast<char>(c);
                i++;
                continue;
            }

            const bool wide = (typeText[i + 1] == L'%');
            if (wide && !TakesWideSuffix(c)) { p.firstBadCode = '%'; return p; }
            const Kind k = KindOf(c, wide);

            if (k == Kind::Unknown)
            {
                p.firstBadCode = static_cast<char>(c);
                return p;   // decline the whole function; do not guess
            }
            if (k == Kind::AsyncHandle) p.async = true;

            // typetext: one code per parameter whatever its slot width, cut with ",..." rather than short.
            if (!sigCut)
            {
                const int codeLen = wide ? 2 : 1;
                if (sigLen + 1 + codeLen + 4 >= static_cast<int>(sizeof(p.signature))) { sigPut(",..."); sigCut = true; }
                else
                {
                    const char code[3] = { static_cast<char>(c), wide ? '%' : '\0', '\0' };
                    if (p.paramCount) sigPut(",");
                    sigPut(code);
                }
            }
            ++p.paramCount;

            // One type code, one ABI slot -- or THREE for O: u16*/i32* rows,
            // then cols, then the double array. Only the first carries the
            // code, and only it renders.
            const int width = (k == Kind::ArrayTriple) ? 3 : 1;
            if (abi + width > Plan::kMaxAbiSlots) { p.firstBadCode = '+'; return p; }
            for (int t = 0; t < width; t++, abi++)
            {
                // Counted even when not described: the thunk forwards every slot.
                if (abi >= Plan::kMaxDescribed) continue;
                Slot& s = p.slots[abi];
                s.kind = k;
                s.abiIndex = abi;
                s.tripleHead = (width == 3 && t == 0);
                s.code[0] = static_cast<char>(c);
                if (wide) s.code[1] = '%';
            }

            i += wide ? 2 : 1;
        }

        p.slotCount = abi;
        p.describedCount = (abi < Plan::kMaxDescribed) ? abi : Plan::kMaxDescribed;

        // rettype: the return code with the registration flags after it ("Q$").
        {
            int r = 0;
            for (const char* rc = p.returnCode; *rc && r < 3; ++rc) p.returnText[r++] = *rc;
            for (int f = 0; f < nflags && r < static_cast<int>(sizeof(p.returnText)) - 1; ++f) p.returnText[r++] = flags[f];
            p.returnText[r] = 0;
        }
        p.ok = true;
        return p;
    }
}
