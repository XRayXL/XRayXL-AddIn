#pragma once
#include "framedwriter.h"

// One value as JSON (docs/TraceRowModel.md). Every value names its type, Double included:
//
//    {"t":"Double","v":1.5}
//    {"t":"Currency","v":"1.5000"}                  exact decimals and 64-bit integers are strings
//    {"t":"Array","elem":"Variant","bounds":[[1,2],[1,3]],"v":[[{"t":"Double","v":11},...],...]}
//    {"t":"Array","elem":"Long","bounds":[[0,3]],"v":[1,2,3,4]}      a typed array's elements are
//                                                                     bare: "elem" names their type
//    {"t":"Object","class":"Range","ptr":"0x1E2...","where":"[Book1]Sheet1!A1","value":{...}}
//    {"t":"Array","elem":"Long","bounds":[[0,99999999]],"omitted":"over the value limit"}
//    [{"slot":1,"type":"Long","value":{"t":"Long","v":5}}]               an argument list
namespace core
{
    class JsonValueWriter final : public FramedWriter
    {
    public:
        explicit JsonValueWriter(TextBuf& out) : FramedWriter(out) {}

        void Double(double v) override;
        void Number(const char* type, const char* text) override;
        void Bool(bool v) override;
        void String(const wchar_t* w, int n, bool cut) override;
        void Error(const char* text) override;
        void Word(const char* w) override;
        void Marker(const char* m) override;

        void BeginObject(const char* cls, std::uint64_t ptr, const char* where) override;
        void EndObject() override;
        void Udt(std::uint64_t ptr) override;
        void BeginReference(const char* kind) override;
        void Area(long long r1, long long c1, long long r2, long long c2) override;
        void EndReference() override;

        void BeginVariant() override;
        void EndVariant() override;

        void BeginArray(const char* elemType, int dims,
                        const long long* lo, const long long* hi) override;
        void BeginLevel() override;
        void EndLevel() override;
        void EndArray() override;
        void Unallocated(const char* elemType) override;

        void BeginArgs() override;
        void BeginArg(int slot, const char* type) override;
        void ArgUnreadable() override;
        void EndArg() override;
        void ArgsNote(int described, int total) override;
        void EndArgs() override;

    private:
        void Sep();
        // A scalar: bare inside a typed array, otherwise {"t":type,"v":json}.
        void Scalar(const char* type, const char* json);
    };

    // `s` as a JSON string literal, quotes included. UTF-8 passes through; a byte that is not
    // part of valid UTF-8 is escaped as the Latin-1 character it would be.
    void AppendJsonString(TextBuf& out, const char* s);
}
