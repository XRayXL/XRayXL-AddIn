#pragma once
#include "jsonvalue.h"
#include "textvalue.h"
#include "tracemodes.h"

#include <atomic>

namespace core
{
    // The format of the file being written, latched when it opens so every value in one file is
    // spelt the same way, whatever the setting does meanwhile.
    inline std::atomic<int> g_activeFormat{ static_cast<int>(modes::Format::Csv) };
    inline modes::Format ActiveFormat()              { return static_cast<modes::Format>(g_activeFormat.load()); }
    inline void          LatchFormat(modes::Format f) { g_activeFormat.store(static_cast<int>(f)); }

    // Calls `fn` with the writer for the active format, writing into `out`.
    template <class Fn>
    void WithValueWriter(TextBuf& out, Fn fn)
    {
        if (ActiveFormat() == modes::Format::Jsonl) { JsonValueWriter w(out); fn(w); }
        else                                        { TextValueWriter w(out); fn(w); }
    }
}
