#pragma once
#include <cstdint>
#include "core/valuewriter.h"

// The one part of the tracer that calls the object model, which runs Excel's own code, so it
// sits behind the OBJECTS setting and is confined to this file. Identity is QueryInterface
// against a known IID, never a resemblance.

namespace vba
{
    // Detailed, named only, or rendered as an address. A class that stops being recognised
    // shows as `described` falling and `namedOnly` rising.
    void ObjectTotals(long long& described, long long& namedOnly, long long& unknown);
    void ResetObjectTotals();

    // `ptr` is an IUnknown*/IDispatch* from a VBA slot:
    //
    //    Worksheet@0x27DB4C21EE0([Book1]Sheet1)
    //    Range@0x27DB4C21EE0([Book1]Sheet1!A1:B2)=Variant[1..2,1..2]{{1,2},{"x",TRUE}}
    //    Collection@0x27DB4C21EE0=Variant[1..2]{1,Nothing}
    //    Dictionary@0x27DB4C21EE0=Variant[0..1,0..1]{{"a",1},{"b",Empty}}   one {key,item} row each
    //
    // False, having written nothing, when the class cannot be named. A COM call that faults
    // costs this description and nothing else.
    bool DescribeObjectDetail(std::uint64_t ptr, core::ValueWriter& w);

    // Where a Range, Worksheet or Workbook is, as its description writes it (`[Book1]Sheet1!A1`),
    // for an event row's `callerref`. False, with `out` empty, for anything else.
    bool ObjectWhere(std::uint64_t ptr, char* out, int cap);
}
