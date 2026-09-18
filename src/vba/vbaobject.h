#pragma once
#include <cstdint>

// Describes a COM object argument or result. The one part of the tracer that calls the object
// model, which runs Excel's own code, so it sits behind the OBJECTS setting and is confined to
// this file.
//
// Identity is QueryInterface against a known IID, never a resemblance. Any failure renders the
// address alone. A range's cell count is checked before Value2 is requested, so a whole-column
// reference never materialises a million-element array inside a calculation.

namespace vba
{
    // For the disarm report: `described` is an object we could detail, `namedOnly` one we could
    // only name, `unknown` one rendered as an address. A class that stops being recognised
    // shows up as `described` falling and `namedOnly` rising.
    void ObjectTotals(long long& described, long long& namedOnly, long long& unknown);
    void ResetObjectTotals();

    // `ptr` is an IUnknown*/IDispatch* taken from a VBA slot. Writes the class name and,
    // separately, whatever detail we can fetch. The caller renders `<Class>@0x<addr><detail>`:
    //
    //    Collection@0x27DB4C21EE0
    //    Worksheet@0x27DB4C21EE0([Book1]Sheet1)
    //    Range@0x27DB4C21EE0([Book1]Sheet1!A1:B2)=Variant[1..2,1..2]{1,2,"x",True}
    //    object@0x27DB4C21EE0                      -- nothing worked out
    //
    // The address stays because it is how one object is followed from an argument to a result.
    // Every path is guarded: a COM call that faults costs this description and nothing else.
    bool DescribeObjectDetail(std::uint64_t ptr, char* cls, int clsCap,
                              char* detail, int detailCap);
}
