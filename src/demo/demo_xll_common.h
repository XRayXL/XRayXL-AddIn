// Shared boilerplate for the XRayXL demo XLLs. Each demo XLL is a separate DLL
// that includes this once; only xlAutoOpen, which registers that XLL's own
// functions, lives in each .cpp.
//
// This is a normal, boring XLL, on purpose: it is the add-in you point XRayXL
// at, not part of XRayXL. Nothing here knows the tracer exists.
#pragma once
#include "plain_xll.h"
#include <cstdio>
#include <cstring>

namespace demo
{
    using namespace plainxll;

    inline thread_local XLOPER12 t_arr[256];   // for an xltypeMulti (array) result

    // A 1-row array {v[0], v[1], ...}, so a modern Excel spills it.
    inline LPXLOPER12 RetRow(const double* v, int n)
    {
        if (n < 1) n = 1; if (n > 256) n = 256;
        for (int i = 0; i < n; i++) { t_arr[i].xltype = xltypeNum; t_arr[i].val.num = v[i]; }
        t_ret.xltype = xltypeMulti;
        t_ret.val.array.rows = 1;
        t_ret.val.array.columns = n;
        t_ret.val.array.lparray = t_arr;
        return &t_ret;
    }
}
