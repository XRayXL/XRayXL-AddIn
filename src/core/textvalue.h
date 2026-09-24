#pragma once
#include "framedwriter.h"

// The trace's text grammar for one value (docs/TraceRowModel.md):
//
//    Long[0..3]{1,2,3,4}
//    Variant[1..2,1..3]{{11,12,13},{21,22,23}}       one brace level per dimension, rows first
//    Variant[0..1]{Integer(1),"x"}                   in a Variant, a number other than Double is named
//    Range@0x1E2...('[Book1]Sheet1'!A1:B1)=Variant[1..1,1..2]{{1,2}}
//    Double[0..99999999]                             no braces: over kMaxValueBytes, shape only
//    a1:Long=5 a2:Variant=Integer(3)                 an argument list
namespace core
{
    class TextValueWriter final : public FramedWriter
    {
    public:
        explicit TextValueWriter(TextBuf& out) : FramedWriter(out) {}

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
        void BeginArg(int slot, const char* type, std::uint64_t address) override;
        void ArgUnreadable() override;
        void EndArg() override;
        void ArgsNote(int described, int total) override;
        void EndArgs() override;

    private:
        void Sep();
    };
}
