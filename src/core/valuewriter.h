#pragma once
#include <cstddef>
#include <cstdint>

// What the decoders know about one value, as a stream of events, so the decoders never depend on
// the output format. Each format is a writer: TextValueWriter for the CSV text grammar in
// docs/TraceRowModel.md, JsonValueWriter for JSON Lines.
namespace core
{
    class ValueWriter
    {
    public:
        // A position to go back to when a decode fails part-way. Opaque to decoders.
        struct Mark { std::size_t len; int frames; int arrays; std::size_t body; bool first; bool pending; };

        // ---- scalars
        virtual void Double(double v) = 0;                           // the default number
        virtual void Number(const char* type, const char* text) = 0; // every other numeric type, exact text
        virtual void Bool(bool v) = 0;
        virtual void String(const wchar_t* w, int n, bool cut) = 0;  // `cut`: the source was longer
        virtual void Error(const char* text) = 0;                    // `#N/A`, or `Error(0x...)` outside Excel's set
        virtual void Word(const char* w) = 0;                        // Empty, Null, Missing, Nothing, AsyncHandle
        virtual void Marker(const char* m) = 0;                      // what is not known: `?`, `?vt17`, `[...]`, `0x...`

        // ---- objects and references
        // `cls` null when the class could not be named; `where` is an address, sheet or book.
        // At most one value between the two is the object's contents, such as a Range's Value2.
        virtual void BeginObject(const char* cls, std::uint64_t ptr, const char* where) = 0;
        virtual void EndObject() = 0;
        virtual void Udt(std::uint64_t ptr) = 0;
        // An XLL reference: `kind` is SRef or Ref, then one Area per rectangle, 1-based.
        virtual void BeginReference(const char* kind) = 0;
        virtual void Area(long long r1, long long c1, long long r2, long long c2) = 0;
        virtual void EndReference() = 0;

        // ---- a Variant's held value sits between these, so its type can be told apart
        virtual void BeginVariant() = 0;
        virtual void EndVariant() = 0;

        // ---- arrays: the header, then one level per dimension, outermost first
        // `elemType` null when the element type could not be named. Bounds in declaration order.
        virtual void BeginArray(const char* elemType, int dims,
                                const long long* lo, const long long* hi) = 0;
        virtual void BeginLevel() = 0;
        virtual void EndLevel() = 0;
        virtual void EndArray() = 0;
        // A dynamic array never allocated: an element type, and no dimensions at all.
        virtual void Unallocated(const char* elemType) = 0;
        // True once nothing more can be kept, so a walk can stop early.
        virtual bool Full() const = 0;

        // ---- an argument list: one entry per slot, each holding at most one value
        virtual void BeginArgs() = 0;
        // `address`: the argument's storage, so a reader can match it to a caller's slot; 0 for none.
        virtual void BeginArg(int slot, const char* type, std::uint64_t address) = 0;
        // An event's parameter, known by its name rather than a slot. `type` null: none is shown.
        virtual void BeginNamedArg(const char* name, const char* type) = 0;
        virtual void ArgUnreadable() = 0;                            // in place of a value
        virtual void EndArg() = 0;
        virtual void ArgsNote(int described, int total) = 0;         // fewer slots described than exist
        virtual void EndArgs() = 0;

        virtual Mark Save() const = 0;
        virtual void Restore(const Mark& m) = 0;

    protected:
        ~ValueWriter() = default;
    };

    // Walks an array row by row, the last index fastest, opening one level per dimension.
    // `leaf(idx)` writes the element at zero-based declaration-order indices `idx`.
    template <class Leaf>
    void WalkLevel(ValueWriter& w, int d, int dims, const std::uint64_t* extent,
                   std::uint64_t* idx, Leaf& leaf)
    {
        w.BeginLevel();
        for (idx[d] = 0; idx[d] < extent[d] && !w.Full(); ++idx[d])
        {
            if (d + 1 < dims) WalkLevel(w, d + 1, dims, extent, idx, leaf);
            else              leaf(idx);
        }
        w.EndLevel();
    }

    template <class Leaf>
    void WalkRowMajor(ValueWriter& w, int dims, const std::uint64_t* extent, Leaf leaf)
    {
        std::uint64_t idx[8] = {};
        if (dims < 1 || dims > 8) return;
        WalkLevel(w, 0, dims, extent, idx, leaf);
    }
}
