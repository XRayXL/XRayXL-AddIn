#pragma once
#include "xlcall.h"
#include "core/ascii.h"
#include "core/text.h"

// An Excel array result: a fixed array of XLOPER12 cells, a pool of counted strings they point at,
// and the grid XLOPER wrapping them. Declare the instance `__declspec(thread)`. The last cell is
// reserved for Scalar(), so a single-cell reply never collides with a grid.
namespace app
{
    template <int Cells, int TextMax>
    struct XlGrid
    {
        XLOPER12 cell[Cells];
        XCHAR    pool[Cells][TextMax];
        XLOPER12 grid;

        // Copied into this cell's pool slot as a counted string, truncated to fit.
        void Str(int i, const wchar_t* text)
        {
            int len = 0; while (text[len] != 0 && len < TextMax - 2) ++len;
            pool[i][0] = static_cast<XCHAR>(len);
            for (int k = 0; k < len; ++k) pool[i][1 + k] = static_cast<XCHAR>(text[k]);
            cell[i].xltype = xltypeStr;
            cell[i].val.str = pool[i];
        }
        // A narrow (char) string cell, widened from UTF-8.
        void StrA(int i, const char* text)
        {
            wchar_t w[TextMax];
            const int len = core::WidenUtf8(text, w, TextMax - 1);
            pool[i][0] = static_cast<XCHAR>(len);
            for (int k = 0; k < len; ++k) pool[i][1 + k] = static_cast<XCHAR>(w[k]);
            cell[i].xltype = xltypeStr;
            cell[i].val.str = pool[i];
        }
        // A number cell, so a sheet can sort and sum it.
        void Num(int i, double v) { cell[i].xltype = xltypeNum; cell[i].val.num = v; }

        // Wrap the first rows*cols cells as a 2-D array result.
        LPXLOPER12 AsGrid(int rows, int cols)
        {
            grid.xltype = xltypeMulti;
            grid.val.array.rows = rows;
            grid.val.array.columns = cols;
            grid.val.array.lparray = cell;
            return &grid;
        }
        // A single-cell string reply, in the reserved last slot.
        LPXLOPER12 Scalar(const wchar_t* text) { Str(Cells - 1, text); return &cell[Cells - 1]; }
    };
}
