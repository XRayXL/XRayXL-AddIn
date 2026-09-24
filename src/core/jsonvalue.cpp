#include "jsonvalue.h"
#include "render.h"

#include <cmath>
#include <cstdio>
#include <cstring>

namespace core
{
    namespace
    {
        const char kHex[] = "0123456789abcdef";

        void AppendU(TextBuf& out, unsigned v)
        {
            const char e[6] = { '\\', 'u', kHex[(v >> 12) & 15], kHex[(v >> 8) & 15],
                                kHex[(v >> 4) & 15], kHex[v & 15] };
            out.Append(e, 6);
        }

        // The escapes JSON requires; false for a character that needs none.
        bool AppendEscape(TextBuf& out, unsigned c)
        {
            switch (c)
            {
            case '"':  out.Append("\\\"", 2); return true;
            case '\\': out.Append("\\\\", 2); return true;
            case '\n': out.Append("\\n", 2);  return true;
            case '\r': out.Append("\\r", 2);  return true;
            case '\t': out.Append("\\t", 2);  return true;
            default:
                if (c < 0x20) { AppendU(out, c); return true; }
                return false;
            }
        }

        // The length of the UTF-8 sequence starting at `s`, or 0 when it is not valid UTF-8.
        int Utf8Length(const unsigned char* s)
        {
            const unsigned c = s[0];
            const int n = c >= 0xF0 && c <= 0xF4 ? 4 : c >= 0xE0 ? 3 : c >= 0xC2 && c <= 0xDF ? 2 : 0;
            for (int i = 1; i < n; ++i) if ((s[i] & 0xC0) != 0x80) return 0;
            return n;
        }

        void AppendUtf8(TextBuf& out, unsigned cp)
        {
            char b[4]; int n;
            if (cp < 0x80)         { b[0] = static_cast<char>(cp); n = 1; }
            else if (cp < 0x800)   { b[0] = static_cast<char>(0xC0 | (cp >> 6));
                                     b[1] = static_cast<char>(0x80 | (cp & 0x3F)); n = 2; }
            else if (cp < 0x10000) { b[0] = static_cast<char>(0xE0 | (cp >> 12));
                                     b[1] = static_cast<char>(0x80 | ((cp >> 6) & 0x3F));
                                     b[2] = static_cast<char>(0x80 | (cp & 0x3F)); n = 3; }
            else                   { b[0] = static_cast<char>(0xF0 | (cp >> 18));
                                     b[1] = static_cast<char>(0x80 | ((cp >> 12) & 0x3F));
                                     b[2] = static_cast<char>(0x80 | ((cp >> 6) & 0x3F));
                                     b[3] = static_cast<char>(0x80 | (cp & 0x3F)); n = 4; }
            out.Append(b, static_cast<std::size_t>(n));
        }

        // UTF-16 as a JSON string literal. A lone surrogate is escaped, since it has no UTF-8.
        void AppendJsonWide(TextBuf& out, const wchar_t* w, int n)
        {
            out.Append('"');
            for (int i = 0; i < n; ++i)
            {
                const unsigned c = static_cast<unsigned>(w[i]);
                if (AppendEscape(out, c)) continue;
                if (c >= 0xD800 && c <= 0xDBFF && i + 1 < n &&
                    w[i + 1] >= 0xDC00 && w[i + 1] <= 0xDFFF)
                {
                    AppendUtf8(out, 0x10000 + ((c - 0xD800) << 10) + (static_cast<unsigned>(w[i + 1]) - 0xDC00));
                    ++i;
                }
                else if (c >= 0xD800 && c <= 0xDFFF) AppendU(out, c);
                else                                  AppendUtf8(out, c);
            }
            out.Append('"');
        }

        // Types whose exact value a double cannot hold are written as strings.
        bool ExactAsString(const char* type)
        {
            static const char* const kExact[] = { "Currency", "Decimal", "LongLong", "UInt64" };
            for (const char* t : kExact)
                if (strcmp(type, t) == 0) return true;
            return false;
        }

        bool IsJsonNumber(const char* s)
        {
            if (*s == '-') ++s;
            if (*s < '0' || *s > '9') return false;
            for (; *s; ++s)
                if (!((*s >= '0' && *s <= '9') || *s == '.' || *s == 'e' || *s == 'E' ||
                      *s == '+' || *s == '-')) return false;
            return true;
        }

        void AppendPtr(TextBuf& out, std::uint64_t ptr)
        {
            char t[32];
            _snprintf_s(t, _TRUNCATE, "\"0x%llX\"", static_cast<unsigned long long>(ptr));
            out.Append(t);
        }
    }

    void AppendJsonString(TextBuf& out, const char* s)
    {
        out.Append('"');
        const unsigned char* p = reinterpret_cast<const unsigned char*>(s ? s : "");
        while (*p)
        {
            if (AppendEscape(out, *p)) { ++p; continue; }
            if (*p < 0x80) { out.Append(static_cast<char>(*p)); ++p; continue; }
            const int n = Utf8Length(p);
            if (n) { out.Append(reinterpret_cast<const char*>(p), static_cast<std::size_t>(n)); p += n; }
            else   { AppendU(out, *p); ++p; }
        }
        out.Append('"');
    }

    // Before a value: a comma between elements, and the key that introduces an object's or an
    // argument's value.
    void JsonValueWriter::Sep()
    {
        const Slot* t = Top();
        if (!t) return;
        switch (t->kind)
        {
        case Frame::Level:
        case Frame::Args:   if (TakeFirst()) m_out.Append(',');                return;
        case Frame::Object:
        case Frame::Arg:    if (!TakeFirst()) m_out.Append(",\"value\":");      return;
        default:            return;
        }
    }

    void JsonValueWriter::Scalar(const char* type, const char* json)
    {
        Sep();
        if (Bare()) { m_out.Append(json); return; }
        m_out.Append("{\"t\":");
        AppendJsonString(m_out, type);
        m_out.Append(",\"v\":");
        m_out.Append(json);
        m_out.Append('}');
    }

    void JsonValueWriter::Double(double v)
    {
        if (std::isnan(v)) { Scalar("Double", "\"NaN\""); return; }
        if (std::isinf(v)) { Scalar("Double", v < 0 ? "\"-Infinity\"" : "\"Infinity\""); return; }
        char t[48];
        FormatDouble(v, t, sizeof t);
        Scalar("Double", t);
    }

    void JsonValueWriter::Number(const char* type, const char* text)
    {
        if (!ExactAsString(type) && IsJsonNumber(text)) { Scalar(type, text); return; }
        char t[96];
        _snprintf_s(t, _TRUNCATE, "\"%s\"", text);
        Scalar(type, t);
    }

    void JsonValueWriter::Bool(bool v) { Scalar("Boolean", v ? "true" : "false"); }

    // A cut string is always an object, so the cut can be said.
    void JsonValueWriter::String(const wchar_t* w, int n, bool cut)
    {
        Sep();
        const bool bare = Bare() && !cut;
        if (!bare) m_out.Append("{\"t\":\"String\",\"v\":");
        AppendJsonWide(m_out, w, n > 0 ? n : 0);
        if (cut)   m_out.Append(",\"cut\":true");
        if (!bare) m_out.Append('}');
    }

    void JsonValueWriter::Error(const char* text)
    {
        Sep();
        const bool bare = Bare();
        if (!bare) m_out.Append("{\"t\":\"Error\",\"v\":");
        AppendJsonString(m_out, text);
        if (!bare) m_out.Append('}');
    }

    void JsonValueWriter::Word(const char* w)
    {
        Sep();
        m_out.Append("{\"t\":");
        AppendJsonString(m_out, w);
        m_out.Append('}');
    }

    void JsonValueWriter::Marker(const char* m)
    {
        Sep();
        m_out.Append("{\"t\":\"Unknown\",\"v\":");
        AppendJsonString(m_out, m);
        m_out.Append('}');
    }

    void JsonValueWriter::BeginObject(const char* cls, std::uint64_t ptr, const char* where)
    {
        Sep();
        m_out.Append("{\"t\":\"Object\"");
        if (cls && cls[0]) { m_out.Append(",\"class\":"); AppendJsonString(m_out, cls); }
        m_out.Append(",\"ptr\":");
        AppendPtr(m_out, ptr);
        if (where && where[0]) { m_out.Append(",\"where\":"); AppendJsonString(m_out, where); }
        Push(Frame::Object);
    }

    void JsonValueWriter::EndObject() { m_out.Append('}'); Pop(); }

    void JsonValueWriter::Udt(std::uint64_t ptr)
    {
        Sep();
        m_out.Append("{\"t\":\"Udt\",\"ptr\":");
        AppendPtr(m_out, ptr);
        m_out.Append('}');
    }

    void JsonValueWriter::BeginReference(const char* kind)
    {
        Sep();
        m_out.Append("{\"t\":");
        AppendJsonString(m_out, kind);
        m_out.Append(",\"areas\":[");
        Push(Frame::Reference);
    }

    void JsonValueWriter::Area(long long r1, long long c1, long long r2, long long c2)
    {
        char t[112];
        _snprintf_s(t, _TRUNCATE, "%s[%lld,%lld,%lld,%lld]", TakeFirst() ? "," : "", r1, c1, r2, c2);
        m_out.Append(t);
    }

    void JsonValueWriter::EndReference() { m_out.Append("]}"); Pop(); }

    // Transparent in JSON: every value names its own type already.
    void JsonValueWriter::BeginVariant() { Sep(); Push(Frame::Variant); }
    void JsonValueWriter::EndVariant()   { Pop(); }

    void JsonValueWriter::BeginArray(const char* elemType, int dims,
                                     const long long* lo, const long long* hi)
    {
        Sep();
        m_out.Append("{\"t\":\"Array\",\"elem\":");
        if (elemType) AppendJsonString(m_out, elemType);
        else          m_out.Append("null");
        m_out.Append(",\"bounds\":[");
        for (int d = 0; d < dims; ++d)
        {
            char t[64];
            _snprintf_s(t, _TRUNCATE, "%s[%lld,%lld]", d ? "," : "", lo[d], hi[d]);
            m_out.Append(t);
        }
        m_out.Append(']');
        OpenArray(elemType && strcmp(elemType, "Variant") != 0);
    }

    void JsonValueWriter::BeginLevel()
    {
        const bool bare = Bare();
        if (OpensOutermostLevel()) m_out.Append(",\"v\":");
        else                       Sep();
        m_out.Append('[');
        Push(Frame::Level, bare);
    }

    void JsonValueWriter::EndLevel() { m_out.Append(']'); Pop(); }

    // An array that did not fit keeps its header, and says why it has no contents.
    void JsonValueWriter::EndArray()
    {
        if (CloseArray()) m_out.Append(",\"omitted\":\"over the value limit\"}");
        else              m_out.Append('}');
    }

    // No dimensions at all: an empty bounds list and no "v".
    void JsonValueWriter::Unallocated(const char* elemType)
    {
        Sep();
        m_out.Append("{\"t\":\"Array\",\"elem\":");
        if (elemType) AppendJsonString(m_out, elemType);
        else          m_out.Append("null");
        m_out.Append(",\"bounds\":[]}");
    }

    void JsonValueWriter::BeginArgs() { m_out.Append('['); Push(Frame::Args); }

    void JsonValueWriter::BeginArg(int slot, const char* type, std::uint64_t address)
    {
        char t[48];
        _snprintf_s(t, _TRUNCATE, "%s{\"slot\":%d,\"type\":", TakeFirst() ? "," : "", slot);
        m_out.Append(t);
        AppendJsonString(m_out, type ? type : "?");
        if (address)
        {
            _snprintf_s(t, _TRUNCATE, ",\"address\":\"0x%llX\"", static_cast<unsigned long long>(address));
            m_out.Append(t);
        }
        Push(Frame::Arg);
    }

    void JsonValueWriter::ArgUnreadable() { m_out.Append(",\"unreadable\":true"); }
    void JsonValueWriter::EndArg()        { m_out.Append('}'); Pop(); }

    void JsonValueWriter::ArgsNote(int described, int total)
    {
        char t[80];
        _snprintf_s(t, _TRUNCATE, "%s{\"described\":%d,\"slots\":%d}", TakeFirst() ? "," : "",
                    described, total);
        m_out.Append(t);
    }

    void JsonValueWriter::EndArgs() { m_out.Append(']'); Pop(); }
}
