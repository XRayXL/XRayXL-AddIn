#pragma once
#include <cstddef>

namespace emit
{

// The row-to-CSV formatter (column order, escaping, header) as a pure transformation, so the
// format can be unit-tested apart from the ring and the drain (docs/TraceRowModel.md).

namespace csv
{
    // CRLF-terminated. `seq` and `input` lead, prefixes added by the writer and producer rather
    // than by Fragment. Change it with the column order in rowcsv.cpp and docs/TraceRowModel.md.
    extern const char* const kHeader;
    // The same with the optional `breaks` column after `trust`, when VBA BREAKPOINTS is on for
    // the file. Every row of that file then has it, empty where it does not apply.
    extern const char* const kHeaderBreaks;

    // Named fields, so neither source counts positions. Pointers, not copies: a row is formatted
    // synchronously from the frame that built its text.
    struct Row
    {
        const char* kind     = "";      // entry | exit
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
        // Exit rows only: how many of `ticks` the tracer itself spent inside the activation.
        const char* tracerticks = "";
        const char* trust    = "";      // exit | end | backstop | flush | async
        // VBA exit rows only, and only in a file opened with the column: how often the call
        // stopped at a breakpoint in the editor.
        const char* breaks   = "";
    };

    // The bytes Fragment will write for `row`, without the NUL. A field is never cut, so this
    // is how the caller sizes its buffer.
    std::size_t FragmentSize(const Row& row, bool breaks = false);

    // Every field escaped and comma-joined, ending in CRLF, without the seq/input prefixes. `out`
    // must hold FragmentSize(row) + 1 bytes. `breaks` adds the optional column, as kHeaderBreaks does.
    std::size_t Fragment(const Row& row, char* out, bool breaks = false);
}
}   // namespace emit
