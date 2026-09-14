#pragma once
#include <cstdint>

// FNV-1a 64. A fingerprint for matching, not a security hash.
namespace core
{
    constexpr std::uint64_t kFnvOffset = 0xcbf29ce484222325ULL;
    constexpr std::uint64_t kFnvPrime  = 0x100000001b3ULL;

    inline std::uint64_t Fnv1aByte(std::uint64_t h, std::uint8_t b)
    {
        return (h ^ b) * kFnvPrime;
    }

    inline std::uint64_t Fnv1aText(const char* text, std::uint64_t h = kFnvOffset)
    {
        for (; text && *text; ++text) h = Fnv1aByte(h, static_cast<std::uint8_t>(*text));
        return h;
    }
}
