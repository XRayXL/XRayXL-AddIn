#pragma once

namespace core
{
    // `[Book1]Sheet1`, quoted exactly where Excel quotes it (`'[has space.xlsx]Sheet1'`).
    void QuoteSheetPrefix(const char* prefix, char* out, int cap);
}
