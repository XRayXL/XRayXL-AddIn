// Finds VBE7's p-code dispatch table and the slots to patch. Derives and verifies; patches
// nothing. Pure static analysis of a module image: no probe executed and no stack walked.
//
// Slot indices are constants: the dispatch table is an interface, and index N means the same
// opcode on every build. Every address is derived.
//
// An Image abstraction, so a harness can read VBE7 as a file through the same derivation the
// add-in runs on the loaded module.
#pragma once
#include <cstdint>
#include <cstddef>
#include <string>
#include <vector>

namespace vba
{
    // Counted, never merged: an instrument that cannot tell "never looked" from
    // "looked and found nothing" reports the second as the first.
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

        // WHY it is not usable, when Ok() is false. Without it the caller could
        // only say "VBE7 is not loaded", which is a lie when the module IS
        // loaded and merely failed to parse -- and sends the user to the VB
        // editor, which cannot help.
        virtual Decline              WhyNotOk() const { return Decline::ModuleNotLoaded; }
    };

    const char* DeclineName(Decline d);

    struct PatchSite
    {
        std::uint32_t slot = 0;         // index into the dispatch table
        std::uint32_t handlerRva = 0;   // what the slot currently holds
        const char*   role = "";        // "bos", "exit" or "end"
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
        // THE SHARED INVALID-OPCODE HANDLER. Slots pointing at it are not instructions, so a walk that LANDS on one is
        // not meeting an unknown opcode: it is already lost, and the step that
        // got there used a WRONG length -- a much sharper signal than "an opcode
        // with no length".
        std::uint32_t invalidHandlerRva = 0;
        std::uint32_t invalidSlots = 0;   // how many point at it
        // The opcode set's fingerprint: for every slot, the index of the lowest slot sharing
        // its handler, hashed. It names no address, so it is the same wherever VBE7 loaded, and
        // it says the table is the opcode set kSigLength describes.
        std::uint64_t partitionHash = 0;
        bool          partitionOk = false;   // ...and it matched the pinned one
        // The `End` slot and whether it verified. `End` fires no exit opcode, and this is the
        // only signal that its frames are dead; a failed check degrades that one feature, the
        // depth and parentage of whatever runs after an `End`, rather than refusing the arm.
        bool          endOk = false;
        int           exitGroups = 0;
        std::uint32_t declines[static_cast<int>(Decline::Count_)] = {};
        std::vector<PatchSite> sites;     // the slots to patch
        std::string   detail;
    };

    // Derive and verify. Never writes to the image.
    SlotSet Derive(const Image& img);

    // One of the procedure-exit opcodes? Exposed so a p-code walk can treat
    // the end of a procedure as a clean stop, not a failure to decode.
    bool IsExitSlot(std::uint32_t slot);

    // THE SAME QUESTION, ASKED PROPERLY: does this opcode END the procedure?
    // Not every member of the exit family does: GoSub `Return` and the pre-exit
    // cleanups (vbaslots.h) do not, and a walk that stops on one reads nothing.
    bool IsProcTerminatorSlot(std::uint32_t slot);

    // Does leaving through this exit mean the LAST argument slot is the
    // function's result rather than a parameter? True for the class/form
    // `[out, retval]` exits.
    bool ExitHasTrailingResultSlot(std::uint32_t slot);

    // Every VBA statement starts with one, which makes it a known-good
    // instruction boundary -- the anchor a walk resynchronises on.
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
