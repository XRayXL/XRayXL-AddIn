#pragma once
#include <cstdint>

// DESCRIBING A COM OBJECT ARGUMENT OR RESULT -- the one part of this tracer that
// TALKS to Excel rather than reading it.
//
// Everything else here is passive: memory reads behind SEH, and the flat C API,
// which is documented for a calculating thread and measured to answer there
// (vbatrace.cpp AskCaller). This calls the OBJECT MODEL, which runs Excel's own
// code. That is why it sits behind the OBJECTS setting and why it is confined to
// this file: the rule everywhere else is that observing must not change what is
// observed, and this is the one place that rule is traded for an answer a user
// cannot otherwise get.
//
// WHAT IT WILL NOT DO:
//   - guess a class. Identity is QueryInterface against an IID CONFIRMED at arm
//     against a known instance, never a resemblance -- not "it has an Address
//     property, so it is probably a Range".
//   - report half an answer. Any failure renders the address, exactly as it did
//     before this existed.
//   - ask for a value it cannot afford. A range's cell count is checked BEFORE
//     Value2 is requested, so a whole-column reference never materialises a
//     million-element array inside a calculation.

namespace vba
{
    // WHAT HAPPENED, for the disarm report. `described` is an object we knew how
    // to detail, `namedOnly` one we could name and no more, `unknown` one we
    // could not even name and rendered as an address.
    //
    // There is NO arm-time step to go with these. QueryInterface is its own
    // validation: an interface id that is wrong, or that Excel changes, simply
    // never matches and the object falls through to being named. These counts
    // are how that stops being invisible -- a class that silently stops being
    // recognised shows up as `described` falling and `namedOnly` rising.
    void ObjectTotals(long long& described, long long& namedOnly, long long& unknown);
    void ResetObjectTotals();

    // `ptr` is an IUnknown*/IDispatch* taken from a VBA slot. On success writes
    // the CLASS name and, separately, whatever DETAIL we know how to fetch --
    // and never the address, because the caller owns that.
    //
    // THE ADDRESS IS NOT OPTIONAL. The rendered form is `<Class>@0x<addr><detail>`,
    // so the class REPLACES the word `object` and adds to what was there:
    //
    //   Collection@0x27DB4C21EE0
    //   Worksheet@0x27DB4C21EE0([Book1]Sheet1)
    //   Range@0x27DB4C21EE0([Book1]Sheet1!A1:B2)=Variant[1..2,1..2]{1,2,"x",True}
    //   object@0x27DB4C21EE0                      -- nothing worked out
    //
    // Dropping it was tried and caught by `object-argument-and-return-are-the-
    // same-address`: the address is how ONE OBJECT IS FOLLOWED from an argument
    // to a result, which no class name can do, and the row model promises it.
    //
    // Every path is guarded: a COM call that faults costs this description and
    // nothing else.
    bool DescribeObjectDetail(std::uint64_t ptr, char* cls, int clsCap,
                              char* detail, int detailCap);
}
