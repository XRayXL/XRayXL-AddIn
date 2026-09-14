#pragma once

namespace emit
{

// THE ROW-TO-CSV FORMATTER -- column order, escaping, header -- as a pure
// transformation touching no Windows, file, lock or global, so the format can be
// unit-tested apart from the ring and the drain. It owns the same contract the
// suites' reader test checks from the outside (docs/TraceRowModel.md).

namespace csv
{
    // One escaped row can be large: the argument and return columns each carry
    // tens of KB of rendered array/variant text, and RFC4180 escaping can nearly
    // double it. The caller supplies a buffer of at least this size -- a
    // per-thread heap block, never the stack.
    constexpr int kFragMax = 256 * 1024;

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
        // WHERE THIS ACTIVATION SAT IN THE CALL CHAIN -- on every row, never empty, each
        // source counting its own frames per thread. Beside `span` because `parent` IS a
        // span: the three call-tree fields read together.
        //
        // They were `depth=` and `parent=` inside `note`. Neither qualifies the
        // `ticks` next to it the way `closed=` does -- they are independent
        // facts about the frame, so they are columns.
        const char* parent   = "";      // the span that called this one, 0 at the top
        const char* depth    = "";      // 1 when no other frame of the same source was open
        const char* thread   = "";
        const char* qpc      = "";
        const char* module   = "";
        const char* function = "";
        const char* proc     = "";
        const char* typetext = "";
        // WHO CALLED IT, AS TWO COLUMNS: the KIND, and the description whose
        // meaning that kind decides. `cell` and `sheet` were two columns that
        // had to agree with each other AND with a third; one description
        // carrying "[Book1]Sheet1!B2:D4" says more (a CSE array formula's whole
        // range does not fit in a cell column) in one field that cannot
        // disagree with itself.
        const char* caller    = "";     // cell | name | toolbar | menu | registerid
                                        // | none | unavailable | not-asked | array | unknown
        const char* callerref = "";     // empty when the kind has no description
        const char* argcount = "";
        const char* args     = "";
        const char* ret      = "";
        const char* rettype  = "";
        // How the activation ended: on every exit row, never empty. XLL exits are always `returned`.
        const char* outcome  = "";      // returned | threw | unwound | handled | abandoned | unhandled
        // THE DURATION, AND WHETHER IT CAN BE BELIEVED. Exit rows only.
        //
        // One number, spelled the same by both sources -- the free-text `note`
        // these replaced had VBA writing `ticks=4541` and the XLL side a bare
        // `604` for the same fact, which only stayed invisible while they were
        // buried in a bag of fields.
        //
        // `trust` names WHAT ENDED THE MEASUREMENT, not a verdict on it, so
        // nothing is thrown away: `exit` and `end` are readings, `backstop` and
        // `flush` are upper bounds, and the mapping is in docs/TraceRowModel.md.
        // `exit` means the same thing on both sides -- the return path fired.
        const char* ticks    = "";      // QPC ticks; empty when none exists
        const char* trust    = "";      // exit | end | backstop | flush | async
    };

    // Escape every field, comma-join them in column order, and end with CRLF --
    // WITHOUT the seq/input prefixes, which the writer/producer prepend. `out`
    // must be at least kFragMax bytes. Returns the number of bytes written.
    int Fragment(const Row& row, char* out);
}
}   // namespace emit
