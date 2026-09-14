#pragma once
#include "xlcall.h"
#include "core/ascii.h"
#include "core/text.h"

// BUILDING AN EXCEL ARRAY RESULT, in one place: a fixed array of XLOPER12
// cells, a parallel pool of pascal-counted strings they point at, and a grid
// XLOPER wrapping them, as a template sized by the caller.
//
// THREAD-LOCAL by construction of the instance. The functions using it are not
// registered thread-safe, so Excel calls them on its main thread; a thread-local
// instance stays correct should one ever be. Declare it `__declspec(thread)`.
//
// The LAST cell is reserved for Scalar(), so a single-cell reply can never
// collide with a grid (which fills from cell 0).
namespace app
{
    template <int Cells, int TextMax>
    struct XlGrid
    {
        XLOPER12 cell[Cells];
        XCHAR    pool[Cells][TextMax];
        XLOPER12 grid;

        // A wide-string cell: the text is copied into this cell's pool slot as a
        // pascal-counted string and the cell points at it. Truncated to fit.
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
        // A NUMBER cell, so a sheet can sort and sum it.
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
        // A single-cell string reply, in the reserved LAST slot.
        LPXLOPER12 Scalar(const wchar_t* text) { Str(Cells - 1, text); return &cell[Cells - 1]; }
    };
}
