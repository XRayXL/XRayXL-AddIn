// Finds VBE7's p-code dispatch table and the slots to patch, by static analysis alone;
// patches nothing. Slot indices are constants (index N is the same opcode on every build);
// every address is derived. The Image abstraction lets a harness run the same derivation on
// VBE7 read as a file.
#pragma once
#include <cstdint>
#include <cstddef>
#include <string>
#include <vector>

namespace vba
{
    // Counted separately, so "never looked" is not reported as "found nothing".
    enum class Decline
    {
        ModuleNotLoaded,
        ImageUnreadable,        // a page of the module faulted while scanning
        HeaderUnreadable,
        NotPe64,
        NoExecutableSection,
        NoCandidateRun,
        NoCandidateAgreedOnBos,
        WrongSlotCount,
        ExitGroupNotUniform,
        BosSlotsDiffer,
        ExitSharesBosHandler,
        Count_
    };

    // A PE addressed in RVA space, whether mapped by the loader or read off
    // disk. Read() must fail rather than fault on a bad range.
    class Image
    {
    public:
        virtual ~Image() = default;
        virtual bool                 Read(std::uint32_t rva, void* dst, std::size_t n) const = 0;
        virtual const std::uint8_t*  Scan(std::uint32_t rva, std::size_t n) const = 0;
        virtual std::uint64_t        Base() const = 0;          // VA this image is based at
        virtual std::uint32_t        SizeOfImage() const = 0;
        virtual std::uint32_t        CodeLoRva() const = 0;
        virtual std::uint32_t        CodeHiRva() const = 0;
        virtual bool                 Ok() const = 0;

        // Why Ok() is false, so a module that failed to parse is not reported as not loaded.
        virtual Decline              WhyNotOk() const { return Decline::ModuleNotLoaded; }
    };

    const char* DeclineName(Decline d);

    struct PatchSite
    {
        std::uint32_t slot = 0;         // index into the dispatch table
        std::uint32_t handlerRva = 0;   // what the slot currently holds
        const char*   role = "";        // "bos", "bosbp", "exit" or "end"
    };

    struct SlotSet
    {
        bool          found = false;      // a table was located
        bool          verified = false;   // ...and every structural check passed
        std::uint32_t tableRva = 0;
        std::uint32_t slots = 0;
        std::uint32_t distinctHandlers = 0;
        std::uint32_t runnerUpSlots = 0;  // longest rival run -- the margin
        std::uint32_t bosHandlerRva = 0;
        // Slots on the shared invalid-opcode handler are not instructions: a walk that lands on
        // one used a wrong length to get there.
        std::uint32_t invalidHandlerRva = 0;
        std::uint32_t invalidSlots = 0;   // how many point at it
        // For every slot, the lowest slot sharing its handler, hashed. Address-free, so it says
        // whether this is the opcode set kSigLength describes.
        std::uint64_t partitionHash = 0;
        bool          partitionOk = false;   // ...and it matched the pinned one
        // `End` fires no exit opcode, so this slot is the only sign its frames are dead. A
        // failure costs depth and parentage after an `End`, not the arm.
        bool          endOk = false;
        // A failure costs the break-in-the-editor count and nothing else.
        bool          stopOk = false;
        // A breakpointed statement never reaches the BoS handler. A failure makes such a call
        // open late and loses the breakpoint count, not the arm.
        bool          bosBpOk = false;
        std::uint32_t bosBpHandlerRva = 0;
        int           exitGroups = 0;
        std::uint32_t declines[static_cast<int>(Decline::Count_)] = {};
        std::vector<PatchSite> sites;     // the slots to patch
        std::string   detail;
    };

    // Derive and verify. Never writes to the image.
    SlotSet Derive(const Image& img);

    // Any member of the exit family, including ones that do not end the procedure.
    bool IsExitSlot(std::uint32_t slot);

    // Does this opcode end the procedure? GoSub `Return` and the pre-exit cleanups (vbaslots.h)
    // do not, and a walk that stopped on one would read nothing more.
    bool IsProcTerminatorSlot(std::uint32_t slot);

    // True for the class/form `[out, retval]` exits, where the last argument slot is the result.
    bool ExitHasTrailingResultSlot(std::uint32_t slot);

    // Every statement starts with one, so it is the anchor a walk resynchronises on. True for
    // the breakpoint form too, which has the same operand and length.
    bool IsBosSlot(std::uint32_t slot);

    // A one-line-per-fact report for the action log and the dialog.
    std::string Describe(const SlotSet& s);



    // VBE7 as loaded in this process. Ok() is false when VBA has not loaded.
    class LoadedVbe7 : public Image
    {
    public:
        LoadedVbe7();
        bool                Read(std::uint32_t rva, void* dst, std::size_t n) const override;
        const std::uint8_t* Scan(std::uint32_t rva, std::size_t n) const override;
        std::uint64_t       Base() const override        { return m_base; }
        std::uint32_t       SizeOfImage() const override  { return m_size; }
        std::uint32_t       CodeLoRva() const override    { return m_codeLo; }
        std::uint32_t       CodeHiRva() const override    { return m_codeHi; }
        bool                Ok() const override           { return m_ok; }
        Decline             WhyNotOk() const override     { return m_why; }
    private:
        const std::uint8_t* m_p = nullptr;
        std::uint64_t m_base = 0;
        std::uint32_t m_size = 0, m_codeLo = 0, m_codeHi = 0;
        bool m_ok = false;
        Decline m_why = Decline::ModuleNotLoaded;
    };

}
