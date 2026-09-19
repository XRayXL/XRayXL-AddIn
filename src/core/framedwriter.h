#pragma once
#include "textbuf.h"
#include "valuewriter.h"

#include <atomic>

namespace core
{
    // Arrays written as their shape alone because they passed kMaxValueBytes. Reset at arm.
    inline std::atomic<long long> g_valuesRefused{ 0 };
    inline long long ValuesRefused()      { return g_valuesRefused.load(); }
    inline void      ResetValuesRefused() { g_valuesRefused.store(0); }

    // The bookkeeping every writer shares: which container is open, whether it has had an
    // element yet, where the outermost array's contents begin, and going back to a Mark. Only
    // the spelling differs between formats.
    class FramedWriter : public ValueWriter
    {
    public:
        bool Full() const override { return m_out.Over(); }

        Mark Save() const override
        {
            Mark m{ m_out.Len(), m_depth, m_arrays, m_body, true, true };
            if (const Slot* t = Top()) { m.first = t->first; m.pending = t->pending; }
            return m;
        }

        void Restore(const Mark& m) override
        {
            m_out.Truncate(m.len);
            m_depth  = m.frames;
            m_arrays = m.arrays;
            m_body   = m.body;
            if (Slot* t = Top()) { t->first = m.first; t->pending = m.pending; }
        }

    protected:
        enum class Frame : unsigned char { Array, Level, Variant, Object, Reference, Args, Arg };
        // `first`: nothing written inside yet. `pending`: an array whose outermost level has not
        // opened. `bare`: inside an array whose element type is named, not Variant.
        struct Slot { Frame kind; bool first; bool pending; bool bare; };
        static constexpr int kMaxFrames = 512;

        explicit FramedWriter(TextBuf& out) : m_out(out) {}
        ~FramedWriter() = default;

        const Slot* Top() const
        { return (m_depth > 0 && m_depth <= kMaxFrames) ? &m_frames[m_depth - 1] : nullptr; }
        Slot* Top()
        { return (m_depth > 0 && m_depth <= kMaxFrames) ? &m_frames[m_depth - 1] : nullptr; }

        // Past kMaxFrames the value cannot be spelt correctly, so it is marked as not fitting.
        void Push(Frame k, bool bare = false)
        {
            if (m_depth >= kMaxFrames) { m_out.over = true; ++m_depth; return; }
            m_frames[m_depth++] = Slot{ k, true, true, bare };
        }
        void Pop() { if (m_depth > 0) --m_depth; }

        // Consumes the first-element state of the innermost container, returning whether an
        // element came before this one.
        bool TakeFirst()
        {
            Slot* t = Top();
            if (!t) return false;
            const bool had = !t->first;
            t->first = false;
            return had;
        }

        // After an array's header: the outermost array's contents start here.
        void OpenArray(bool bare)
        {
            if (m_arrays == 0) m_body = m_out.Len();
            ++m_arrays;
            Push(Frame::Array, bare);
        }

        // True when the outermost array did not fit: its contents are dropped back to the
        // header, and the refusal is counted. Never part of the contents.
        bool CloseArray()
        {
            Pop();
            if (m_arrays > 0) --m_arrays;
            if (m_arrays != 0 || !m_out.Over()) return false;
            m_out.Truncate(m_body);
            g_valuesRefused.fetch_add(1);
            ++m_out.refused;
            return true;
        }

        // The outermost level of an array follows its header; any other is an element.
        bool OpensOutermostLevel()
        {
            Slot* t = Top();
            if (t && t->kind == Frame::Array && t->pending) { t->pending = false; return true; }
            return false;
        }

        bool InVariant() const { const Slot* t = Top(); return t && t->kind == Frame::Variant; }
        bool Bare() const      { const Slot* t = Top(); return t && t->bare; }

        TextBuf&    m_out;
        Slot        m_frames[kMaxFrames];
        int         m_depth  = 0;
        int         m_arrays = 0;     // open arrays
        std::size_t m_body   = 0;     // where the outermost array's contents begin
    };
}
