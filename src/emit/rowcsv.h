#pragma once
#include <cstddef>

namespace emit
{

// THE ROW-TO-CSV FORMATTER -- column order, escaping, header -- as a pure
// transformation touching no Windows, file, lock or global, so the format can be
// unit-tested apart from the ring and the drain. It owns the same contract the
// suites' reader test checks from the outside (docs/TraceRowModel.md).

namespace csv
{
    // THE HEADER LINE, CRLF-terminated. `seq` and `input` lead, both prefixes
    // added by the writer/producer rather than by Fragment. Change it and the
    // column order in rowcsv.cpp -- the one place that knows it -- and
    // docs/TraceRowModel.md change with it.
    extern const char* const kHeader;

    // ONE ROW, NAMED, so neither source counts positions
    // (docs/TraceRowModel.md). Pointers, not copies: a row is formatted
    // synchronously from the frame that built its text, so nothing here
    // outlives the call.
    struct Row
    {
        const char* kind     = "";      // entry | exit | depth-capped
        const char* source   = "";      // XLL | VBA
        const char* span     = "";
        // Where this activation sat in the call chain: on every row, never empty, each source
        // counting its own frames per thread.
        const char* parent   = "";      // the span that called this one, 0 at the top
        const char* depth    = "";      // 1 when no other frame of the same source was open
        const char* thread   = "";
        const char* qpc      = "";
        const char* module   = "";
        const char* function = "";
        const char* proc     = "";
        const char* typetext = "";
        // Who called it, as two columns: the kind, and the description whose meaning that kind
        // decides. One description holds a whole range such as "[Book1]Sheet1!B2:D4".
        const char* caller    = "";     // cell | name | toolbar | menu | registerid
                                        // | none | unavailable | not-asked | array | unknown
        const char* callerref = "";     // empty when the kind has no description
        const char* argcount = "";
        const char* args     = "";
        const char* ret      = "";
        const char* rettype  = "";
        // How the activation ended: on every exit row, never empty. XLL exits are always `returned`.
        const char* outcome  = "";      // returned | threw | unwound | handled | abandoned | unhandled
        // The duration, on exit rows only, spelled the same by both sources. `trust` names what
        // ended the measurement, not a verdict on it: `exit` and `end` are readings, `backstop`
        // and `flush` are upper bounds (docs/TraceRowModel.md).
        const char* ticks    = "";      // QPC ticks; empty when none exists
        const char* trust    = "";      // exit | end | backstop | flush | async
    };

    // The bytes Fragment will write for `row`, without the NUL. A field is never cut, so this
    // is how the caller sizes its buffer.
    std::size_t FragmentSize(const Row& row);

    // Escape every field, comma-join them in column order, and end with CRLF --
    // WITHOUT the seq/input prefixes, which the writer/producer prepend. `out`
    // must hold FragmentSize(row) + 1 bytes. Returns the number of bytes written.
    std::size_t Fragment(const Row& row, char* out);
}
}   // namespace emit
