#pragma once
#include "rowcsv.h"
#include "core/textbuf.h"

namespace emit
{
// THE ROW-TO-JSON FORMATTER, for a JSON Lines trace: one object a line, its keys the CSV
// columns in the same order (docs/TraceRowModel.md). `args` and `ret` are already JSON, written
// by core::JsonValueWriter; every other field is a string, or a number where the column only
// ever holds an integer. An empty field has no key.
namespace json
{
    // Appends the row's keys after the writer's `{"seq":N,` and the producer's `"input":N,`,
    // closing the object and the line.
    void AppendRow(const csv::Row& row, core::TextBuf& out);
}
}   // namespace emit
