#include "textvalue.h"
#include "render.h"

#include <cmath>
#include <cstdio>

namespace core
{
    // Before a value: a comma between elements, a space between arguments, the `=` that
    // introduces an object's contents; nothing elsewhere.
    void TextValueWriter::Sep()
    {
        const Slot* t = Top();
        if (!t) return;
        switch (t->kind)
        {
        case Frame::Level:  if (TakeFirst()) m_out.Append(',');   return;
        case Frame::Object: if (!TakeFirst()) m_out.Append('=');  return;
        default:            return;
        }
    }

    void TextValueWriter::Double(double v)
    {
        Sep();
        if (std::isnan(v)) { m_out.Append("NaN"); return; }
        if (std::isinf(v)) { m_out.Append(v < 0 ? "-Infinity" : "Infinity"); return; }
        char t[48];
        FormatDouble(v, t, sizeof t);
        m_out.Append(t);
    }

    // Inside a Variant a bare number means Double, so every other type is named.
    void TextValueWriter::Number(const char* type, const char* text)
    {
        const bool tag = InVariant() && type && type[0];
        Sep();
        if (tag) { m_out.Append(type); m_out.Append('('); }
        m_out.Append(text);
        if (tag) m_out.Append(')');
    }

    void TextValueWriter::Bool(bool v)            { Sep(); m_out.Append(v ? "TRUE" : "FALSE"); }
    void TextValueWriter::Error(const char* text) { Sep(); m_out.Append(text); }
    void TextValueWriter::Word(const char* w)     { Sep(); m_out.Append(w); }
    void TextValueWriter::Marker(const char* m)   { Sep(); m_out.Append(m); }

    void TextValueWriter::String(const wchar_t* w, int n, bool cut)
    {
        Sep();
        // Six bytes a character at worst, then the quotes, "..." and the NUL.
        const std::size_t room = static_cast<std::size_t>(n > 0 ? n : 0) * 6 + 16;
        char* d = m_out.Reserve(room);
        if (!d) return;
        if (!RenderQuoted(w, n, cut, d, static_cast<int>(room + 1))) return;
        m_out.Commit(strlen(d));
    }

    void TextValueWriter::BeginObject(const char* cls, std::uint64_t ptr, const char* where)
    {
        Sep();
        char t[160];
        _snprintf_s(t, _TRUNCATE, "%s@0x%llX", cls && cls[0] ? cls : "object",
                    static_cast<unsigned long long>(ptr));
        m_out.Append(t);
        if (where && where[0]) { m_out.Append('('); m_out.Append(where); m_out.Append(')'); }
        Push(Frame::Object);
    }

    void TextValueWriter::EndObject() { Pop(); }

    void TextValueWriter::Udt(std::uint64_t ptr)
    {
        Sep();
        char t[32];
        _snprintf_s(t, _TRUNCATE, "udt@0x%llX", static_cast<unsigned long long>(ptr));
        m_out.Append(t);
    }

    void TextValueWriter::BeginReference(const char* kind)
    {
        Sep();
        m_out.Append(kind);
        m_out.Append('(');
        Push(Frame::Reference);
    }

    void TextValueWriter::Area(long long r1, long long c1, long long r2, long long c2)
    {
        if (TakeFirst()) m_out.Append(',');
        char t[96];
        _snprintf_s(t, _TRUNCATE, "R%lldC%lld:R%lldC%lld", r1, c1, r2, c2);
        m_out.Append(t);
    }

    void TextValueWriter::EndReference() { m_out.Append(')'); Pop(); }

    void TextValueWriter::BeginVariant() { Sep(); Push(Frame::Variant); }
    void TextValueWriter::EndVariant()   { Pop(); }

    // Always both bounds, because Option Base decides what a bare count would mean.
    void TextValueWriter::BeginArray(const char* elemType, int dims,
                                     const long long* lo, const long long* hi)
    {
        Sep();
        m_out.Append(elemType ? elemType : "?");
        m_out.Append('[');
        for (int d = 0; d < dims; ++d)
        {
            char t[64];
            _snprintf_s(t, _TRUNCATE, "%s%lld..%lld", d ? "," : "", lo[d], hi[d]);
            m_out.Append(t);
        }
        m_out.Append(']');
        OpenArray(false);
    }

    void TextValueWriter::BeginLevel()
    {
        if (!OpensOutermostLevel()) Sep();
        m_out.Append('{');
        Push(Frame::Level);
    }

    void TextValueWriter::EndLevel() { m_out.Append('}'); Pop(); }

    // An array that did not fit keeps its header and loses its braces.
    void TextValueWriter::EndArray() { CloseArray(); }

    // `Long()`: the type an unallocated array has, and nothing else, as VBA's TypeName says.
    void TextValueWriter::Unallocated(const char* elemType)
    {
        Sep();
        m_out.Append(elemType ? elemType : "?");
        m_out.Append("()");
    }

    void TextValueWriter::BeginArgs() { Push(Frame::Args); }

    void TextValueWriter::BeginArg(int slot, const char* type, std::uint64_t address)
    {
        char at[24] = "";
        if (address) _snprintf_s(at, _TRUNCATE, "@0x%llX", static_cast<unsigned long long>(address));
        char t[160];
        _snprintf_s(t, _TRUNCATE, "%sa%d:%s%s=", TakeFirst() ? " " : "", slot, type ? type : "?", at);
        m_out.Append(t);
        Push(Frame::Arg);
    }

    void TextValueWriter::BeginNamedArg(const char* name, const char* type)
    {
        char t[160];
        if (type) _snprintf_s(t, _TRUNCATE, "%s%s:%s=", TakeFirst() ? " " : "", name, type);
        else      _snprintf_s(t, _TRUNCATE, "%s%s=", TakeFirst() ? " " : "", name);
        m_out.Append(t);
        Push(Frame::Arg);
    }

    void TextValueWriter::ArgUnreadable() { m_out.Append("<unreadable>"); }
    void TextValueWriter::EndArg()        { Pop(); }

    void TextValueWriter::ArgsNote(int described, int total)
    {
        char t[64];
        _snprintf_s(t, _TRUNCATE, "%s...(%d of %d slots described)", TakeFirst() ? " " : "",
                    described, total);
        m_out.Append(t);
    }

    void TextValueWriter::EndArgs() { Pop(); }
}
