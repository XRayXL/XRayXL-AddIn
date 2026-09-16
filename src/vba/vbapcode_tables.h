#pragma once
#include <cstdint>

// THE PINNED TABLES OF THE P-CODE WALK -- data, no code. The typed load
// opcodes, the instruction lengths, the variable-length forms and the
// corpus-walked bitmap: every entry measured, each with its provenance
// beside it. Included by vbapcode.cpp alone; the offline tools read the
// tables from this file by name.
namespace vba
{
    namespace
    {
        // THE TYPED ARGUMENT-LOAD OPCODES, each measured by declaring one
        // parameter of that VBA type and reading the opcode the body emits.
        // [measured: frame-types]
        //
        // These are SLOT INDICES, not addresses, and slot indices are portable
        // -- the same property the BoS and exit slots rely on, across 41 builds.
        //
        // 656..669 is one contiguous per-type block; the ByRef loads at 736..751
        // are a second block that reproduces it offset for offset. Each entry was
        // still MEASURED rather than predicted from the mirror.
        //
        // THIS TABLE WAS ONCE EXTENDED ONLY WHEN SOMEBODY HAPPENED TO WRITE the
        // procedure that used a missing opcode. It is now filled from the mined
        // corpus, which reaches shapes nobody would think to write.
        struct TypeOp { std::uint32_t op; const char* name; };
        constexpr TypeOp kTypeOps[] = {
            { 656,  "Byte"     }, { 657,  "Integer"  },   // Integer AND Boolean
            { 658,  "Long"     }, { 661,  "Currency" },
            { 663,  "String"   }, { 664,  "Object"   },
            { 667,  "LongLong" }, { 668,  "Single"   },
            { 669,  "Double"   },                          // Double AND Date
            { 1477, "Variant"  },
            // 788 is a reference to a String but NOT the family member: 788-736
            // is 52, where every other ByRef load sits within 13 of the base. It
            // and 743 are both listed, neither inferred from the other.
            // (770 is the ByRef Long STORE, 738+32, and belongs in the store
            // relation rather than here.)
            { 788,  "String&"  },   // ByRef String -- a real load, see below
            // THE BY-REFERENCE SCALAR FAMILY. Offsets from 736 reproduce the
            // ByVal family exactly -- Byte +0, Integer +1, Currency +5, Object
            // +8, Single +12, Double +13 -- and both alias pairs come out right
            // (Boolean IS I2, a Date IS a Double).
            { 736,  "Byte&"    }, { 737,  "Integer&" },   // Integer AND Boolean
            { 738,  "Long&"    },
            { 741,  "Currency&"}, { 742,  "Variant&" },
            { 743,  "String&"  },
            { 744,  "Object&"  }, { 748,  "Single&"  },
            // 750 is a reference to a *specific* class -- a VBA class module or
            // an imported COM class -- where 744 is the generic `As Object`. Its
            // length is the family's 6; shipped_consensus says 2, the same trap
            // as 788 and 775.
            { 750,  "Object&"  },
            { 749,  "Double&"  },                          // Double AND Date
            // 743 is the family's String&, at 736+7 as ByVal String is 656+7.
            //
            // A user-defined Type is always passed by reference and the slot
            // points straight at the record. 1056..1071 is one family (all
            // operand-8): a MEMBER access through a UDT reference, the member's
            // own type selecting the slot.
            { 1058, "Udt&"     }, { 1069, "Udt&" }, { 1071, "Udt&" },
            // 1063 is inside the member-access family: a String member read
            // through a UDT reference, so the PARAMETER is a `Udt&`. 751 looks
            // like it belongs here and does not -- see PcodeCarriesNoType.
            { 1063, "Udt&" },
            // THE UDT-MEMBER FAMILY IS TYPED BY THE MEMBER, and a `Type`
            // carrying one member of every type reaches all of it:
            //
            //   1056 Byte   1057 Integer/Boolean   1058 Long      1061 Currency
            //   1062 Variant  1063 String   1064 Object   1067 LongLong
            //   1068 Single   1069/1071 (already named)
            //
            // -- offsets from 1056 reproducing the ByVal family's.
            //
            // ALL OF THEM NAME THE PARAMETER `Udt&`, NOT the member: the slot
            // holds a reference to the record, and which member the instruction
            // reads is incidental to the PARAMETER's declared type. Naming them
            // by member would put `(Byte)` on a `ByRef p As TBig`.
            //
            // The matching stores are each exactly load+32, so
            // `PcodeStoreTypeName` resolves them through these rows -- which is
            // the point of the store relation, and a check on it.
            { 1056, "Udt&"     }, { 1057, "Udt&"     }, { 1061, "Udt&" },
            { 1062, "Udt&"     }, { 1064, "Udt&"     }, { 1067, "Udt&" },
            { 1068, "Udt&"     },
            // 1059/1060 are this family's COERCING Single and Double, at the
            // same +3/+4 offsets the ByVal family puts 659/660 at, and reached
            // the same way -- across a Variant-ish boundary. A plain read of the
            // same members emits 1068/1069.
            //
            // 1108 IS THE STRING-MEMBER STORE and an outlier by the rule the
            // ByVal side already records: a String store never sits at load+32,
            // so the store relation cannot reach it and it is listed.
            // [measured: 534 cells]
            { 1059, "Udt&"     }, { 1060, "Udt&"     }, { 1108, "Udt&" },
            // 786 IS THE ByRef CLASS STORE (`Set p = Nothing` on a
            // `ByRef p As CThing`) and cannot come from the store relation:
            // 786-32 = 754, which is unimplemented. Same outlier shape as 788.
            { 786,  "Object&"  },
            // 660/740 ARE DOUBLE at offset 4 of their families -- the coercing
            // pair, appearing ONLY when the value crosses an object member
            // (`c.V = p`, `c.M(p)`); a plain read of the same parameter emits
            // 669/749. What is measured is the TYPE the slot carries, which is
            // what it has to name.
            // [measured: objprop and objmethod]
            { 660,  "Double"   }, { 740,  "Double&"  },
            // 659/739 are the same coercing pair one offset below, appearing
            // only through `Debug.Print p`. A Single is four
            // bytes, and `mov eax,[rax]` is exactly how a float's bit pattern
            // moves -- zero-extending, since sign-extending a float is
            // meaningless. [measured: debugprint]
            { 659,  "Single"   }, { 739,  "Single&"  },
            // AN 8-BYTE BY-REFERENCE LOAD THAT DOES NOT DETERMINE THE TYPE: 747
            // fires for `ByRef v As LongLong`, `As LongPtr` AND `ByRef a() As
            // ...`, so it carries a WIDTH and asserts no type. The "&" still
            // makes the renderer follow the pointer, and the self-validating
            // decoders decide.
            { 747,  "Ref&" },
        };

        // THE VERIFIED LENGTHS OF THE OPCODES A SIGNATURE DEPENDS ON.
        //
        // An instruction's length is a property of the OPCODE SET, not of the
        // build -- 42 VBE7 builds, 2012 to 2026, zero conflicts
        // -- so it is a constant for the same reason the slot indices in
        // vbaderive.h are: a dispatch table is an interface. The per-slot build
        // counts are regenerated offline from the binaries.
        //
        // FOUR RULES DECIDE WHAT GETS AN ENTRY.
        //
        //   1. Unanimous in `candidate_consensus` across the corpus -- NEVER
        //      `shipped_consensus`, which is wrong for 788, 750 and 775.
        //   2. NOT a desync artefact. An opcode that stopped many walks but
        //      derives on NO build (op8, 710, and the 0/1/2/4/16/56 run) is
        //      operand bytes read as an opcode after the walk lost alignment.
        //      Pinning one would let a misaligned walk carry on instead of
        //      resynchronising -- the one failure the resync exists to prevent.
        //   3. UNANIMOUS IS NOT PROOF. A length can be unanimously derived and
        //      unanimously WRONG: for a CALL, the `movzx rax, word [rsi]` the
        //      handler reaches is the CALLEE'S first opcode, so a walker
        //      measuring "how far had rsi moved before the next fetch" measures
        //      the wrong thing. 1309, 500, 498 and 1311 all derive as 2 and are
        //      6; only the emitted p-code says so.
        //   4. Exit opcodes are absent: the walk steps over one by the offset the
        //      statement declares, so their length is never used.
        //
        // WHERE THE ENTRIES CAME FROM, weakest evidence last: traced procedures;
        // the shape fuzzer, whose oracle is "did the tracer understand every
        // opcode it met", so it needs no known answer; the handlers themselves, on 2
        // builds rather than 42, so those rows are marked weaker; and statement
        // segment arithmetic -- a statement opcode's operand is the offset of the
        // next statement, so a segment's last unpinned length falls out by subtraction.
        struct SigLength { std::uint16_t slot; std::uint8_t len; };
        constexpr SigLength kSigLength[] = {
            { 615, 6 }, { 1645, 6 },                    // beginning-of-statement
            { 656, 6 }, { 657, 6 }, { 658, 6 },         // Byte, Integer, Long
            { 661, 6 }, { 663, 6 }, { 664, 6 },         // Currency, String, Object
            { 667, 6 }, { 668, 6 }, { 669, 6 },         // LongLong, Single, Double
            { 736, 6 }, { 737, 6 }, { 741, 6 },         // ByRef Byte, Integer, Currency
            { 744, 6 }, { 748, 6 }, { 749, 6 },         // ByRef Object, Single, Double
            { 770, 6 }, { 788, 6 },                     // ByRef Long, ByRef String
            { 747, 6 },                                 // untyped 8-byte reference
            { 750, 6 },                                 // ByRef a named class
            { 742, 10 }, { 1477, 10 },                  // Variant&, Variant
            { 1058, 10 }, { 1069, 10 }, { 1071, 10 },   // UDT member access

            // On the path to a LOAD, but not loads themselves.
            { 689, 6 }, { 690, 6 }, { 699, 6 },
            { 701, 6 }, { 708, 6 }, { 1520, 6 },

            // On the path to a STORE. A resync consumes the rest of a statement,
            // so without these `p = 9.75` never reached its store.
            { 492,  2 }, { 671, 6 }, { 743,  6 },
            { 1525, 10 }, { 1527, 4 }, { 1528, 8 }, { 1655, 2 },

            // The ByRef STORES (load + 32), which pinning the group above made
            // the last instruction the walk reached.
            { 1311, 6 },
            { 769,  6 }, { 775,  6 }, { 781,  6 },

            // From the UDT-member family; it blocked the write-only String
            // assigned from a variable.
            { 1063, 10 },

            // On the path to a load INSIDE A CALL: `f = p * g(q)` and
            // `p = Application.Run("X", q)` left their first parameter untyped
            // because the call setup stopped the walk and the resync ate the
            // load with it.
            { 497,  6 }, { 718,  6 }, { 1534, 8 },
            { 63,   2 }, { 246,  2 }, { 991, 4 }, 

            // 1090 IS THE MODULE-LEVEL STORE (`gSink = 1` where gSink is a
            // module variable), among the most ordinary statements in VBA: it
            // alone stopped 15 of 16 walks, resynchronising the walk out of
            // nearly every real procedure it met.
            // [measured]
            { 1090, 10 }, { 660, 6 },

            // FROM THE SHAPE FUZZER: opcodes that stopped at least two walks in a
            // run and are unanimous across all 42 builds. Those whose length derives
            // on only some builds are left unpinned: that is the cross-version
            // fragility this project avoids.
            // [measured: seeds 3000-3009]
            { 36, 2 }, { 39, 2 }, { 49, 2 }, { 50, 2 }, { 197, 2 },
            { 198, 2 }, { 267, 2 }, { 281, 2 }, { 282, 2 }, { 283, 2 },
            { 294, 2 }, { 332, 2 }, { 356, 2 }, { 357, 2 }, { 388, 2 },
            { 400, 2 }, { 401, 2 }, { 403, 2 }, { 404, 2 }, { 405, 2 },
            { 411, 2 }, { 424, 2 }, { 425, 2 }, { 426, 2 }, { 429, 2 },
            { 435, 2 }, { 436, 2 }, { 437, 2 }, { 539, 2 }, { 546, 2 },
            { 589, 2 }, { 597, 6 }, { 659, 6 }, { 662, 6 }, { 670, 6 },
            { 672, 6 }, { 688, 6 }, { 693, 6 }, { 700, 6 }, { 711, 6 },
            { 738, 6 }, { 751, 6 }, { 951, 8 }, { 1032, 6 }, { 1070, 10 },
            { 1089, 10 }, { 1101, 10 }, { 1111, 4 }, { 1114, 4 }, { 1265, 6 },
            { 1266, 6 }, { 1271, 6 }, { 1303, 6 }, { 1415, 10 }, { 1472, 4 },
            { 1473, 10 }, { 1517, 4 }, { 1524, 10 }, { 1634, 2 }, { 1648, 2 },
            { 1649, 2 }, { 1650, 2 }, { 1652, 2 },

            // A later shape-fuzz batch, each unanimous across all 42 corpus builds.
            { 1, 2 }, { 27, 2 }, { 37, 2 }, { 38, 2 }, { 51, 2 }, { 147, 2 },
            { 148, 2 }, { 150, 2 }, { 195, 2 }, { 220, 2 }, { 232, 2 },
            { 279, 2 }, { 299, 2 }, { 352, 2 }, { 354, 2 }, { 363, 2 },
            { 441, 2 }, { 443, 2 }, { 485, 2 }, { 537, 2 },
            { 621, 2 }, { 622, 2 }, { 878, 6 }, { 909, 6 },
            { 1609, 6 }, { 1630, 6 }, { 1651, 2 }, { 1687, 2 },

            // FROM THE HANDLERS THEMSELVES, not the corpus CSV. These
            // read `derivedOn=0/42` in the corpus CSV, which looked like "VBE7
            // does not say". It said: the offline analysis reads
            // the handler under two extra rules --
            //
            //   R1  an explicit `add rsi,N` SURVIVES losing the thread. 696 and
            //       1441 state their length outright and then jump into a shared
            //       tail or make a call; a walk insisting on reaching a
            //       `movzx rax, word [rsi]` fetch throws that answer away.
            //   R2  a load INTO rsi FROM [rsi+k] is a BRANCH, not an advance.
            //       710 is `mov esi,[rsi]` + `add rsi,base`, which REPLACES the
            //       p-code pointer -- but it still occupies the 4 bytes of target
            //       it read, which is what a LINEAR walk needs.
            //
            // A REJECTED RULE, RECORDED: "treat `call` as a stop" sounds safer,
            // but measured it answers 90 new slots and LOSES 184, because the
            // idiom is `call helper; ...; movzx rax,[rsi]` -- the fetch is after
            // the call, not inside it. Recorded so it is not
            // tried again.
            //
            // WEAKER EVIDENCE THAN THE ROWS ABOVE, AND SAID SO. Those are 42/42
            // from the corpus CSV. These are 2/2: measured on the two
            // builds at hand, not across the corpus.
            // What IS 42-build-grade is the pair of rules, which contradict none
            // of the ~800 lengths the walk derives on either build.
            // [measured: on 7.1.10.33 and 7.1.11.58; every one met repeatedly by the shape fuzzer]
            { 696,  6 },  { 707,  6 },  { 710,  6 },  { 959,  6 },
            { 1468, 10 },  { 1471, 10 },
            { 1441, 12 },  { 1442, 12 },
            { 1096, 10 }, { 1107, 10 }, { 1445, 10 },

            // 1309 IS THE PROCEDURE CALL AND IT IS 6, not the 2 that both
            // handler walkers derive unanimously (cand=2, shipped=2,
            // derivedOn=42/42, and the offline length derivation agrees on both
            // builds). The bytes settle it --
            // `RiskWeighted = exposure * WeightFor(rating)` compiles to:
            //
            //   6702 0600 0000   615@6     beginning of statement
            //   9202 1000 0000   658@16    load Long   <- rating
            //   1D05 0200 0800   1309      CALL WeightFor, 4-byte operand
            //   BD02 F0FF FFFF   701@-16   store Double to a temp
            //   9D02 0800 0000   669@8     load Double <- exposure
            //   9D02 F0FF FFFF   669@-16   load the temp back
            //   F600             246       multiply
            //   BD02 F8FF FFFF   701@-8    store to the return slot
            //   6702 0000 0000   615@0
            //   7302             627       exit, Double
            //
            // At length 6 every following opcode decodes as that statement; at 2
            // the walk lands on `0200` mid-operand, reports a phantom "opcode 2",
            // resynchronises past the rest of the statement and loses the load of
            // `exposure`. [measured: XRAYXL_DIAG dump, RiskWeighted;
            //  defect: call-length-derived-from-the-callee]
            //
            // op500 and op498 are the same defect at the member dispatch, and
            // op1311 the third instance: each pinned 2, each really 6, each
            // found from the compiler's OUTPUT rather than the handler. 1311
            // alone takes the corpus from 443 to 527 procedures parsed (70% ->
            // 83%), breaking none.
            // [measured: the ClassParam dump]
            //
            // op406 = 2, and op1497 stays 10. The corpus ranks {1497=10,406=2}
            // and {1497=10,406=8} EQUAL, so it cannot choose; the procedure
            // decides. `Sub P(ByVal r As Range) / Dim z As Long / z = r.Row` must
            // STORE a Long into z, and only 406=2 leaves op690 (the Long store)
            // in the stream -- with 8 the assignment has no store at all, which
            // no compiler emits.
            { 498, 6 },
            { 406, 2 },
            // op600 and op619 by SEGMENT ARITHMETIC (13 and 3 independent
            // segments, unanimous), confirmed by the hold-out solver.
            // op200 = 6: a 4-byte operand, confirmed by its handler and by the dispatch tail.
            { 200, 6 },
            { 600, 2 }, { 619, 2 },
            // op620 (GoSub Return) = 2: no operand, per the dispatch tail. The handler reloads
            // RSI before fetching, so reading the handler alone cannot show this.
            { 620, 2 },
            // THE CALL FAMILIES AT 6, FROM THE HANDLERS THEMSELVES. At dispatch RSI
            // points at the operand, and a call handler reads its operand words,
            // `add rsi,4`, SAVES rsi to a frame slot, calls, restores it, and only
            // then fetches the next opcode -- so the length is 2 + the advance made
            // before the save, and a walker that stops at the first [rsi] read
            // answers 2 for an instruction that is 6. Read from the handler bytes
            // on 7.1.11.58, validated on the nine calls the corpus had already
            // proved (all read 6), and on the corpus itself: with 1279 at 6 the
            // thirty-four parses that broke immediately after it all close, and
            // ImpAdCall and ThisVCall hold on the unseen half. The other VCall
            // slots never occur in any traced code; they share the handler shape
            // and the name family of the two (498, 500) the corpus proved, and 2
            // is known wrong for that family, so 6 is the lower-risk value -- held
            // at lower confidence until code that emits them is traced.
            // The through-call rule is NOT trusted outside the call families: on AddVar and
            // Ary1LdRfVarg it contradicts corpus-validated pins and applying it
            // there breaks 56 procedures. [measured: hold-out check]
            { 500, 6 },
            { 1309, 6 },

            // 1606 IS 4, not the 2 the corpus derives. The clean-walk counter
            // flagged it as a length the walk USED and could not continue past;
            // of 2/4/6/8 only 4 improved anything, so the answer is unique rather
            // than a tie.
            //
            // Four other suspects from the same ranking were tried and LEFT
            // ALONE, which is the discipline this loop enforces: op283 and
            // op1271 changed nothing, op1111 RAISED desyncs, and op1266's
            // candidates tie exactly -- unpinned beats wrongly pinned.
            // [measured: seed 3000]
            { 1606, 4 },

            // EVERY LENGTH THE HANDLER WALK DERIVES, not only the ones the fuzzer had
            // happened to meet -- which left 725 cross-build-consistent slots
            // unpinned on the (true, but only about COST) principle that an
            // opcode never met costs nothing unlengthed.
            //
            // Clean walks 19 -> 23 of 53, resyncs 89 -> 80, `)~` signatures
            // 34 -> 30, and DESYNCS UNCHANGED AT 1 -- the check that matters,
            // because a wrong length shows up as a desync before anywhere else.
            // Identical on both locally available builds, and contradicting no
            // pinned length except 1309 and 1606, corrected above from the
            // p-code. [measured: the shape fuzzer]
            { 2, 2 }, { 3, 2 }, { 12, 2 }, { 13, 2 }, { 14, 2 }, { 15, 2 },
            { 24, 2 }, { 25, 2 }, { 26, 2 }, { 48, 2 }, { 60, 2 }, { 61, 2 },
            { 62, 2 }, { 64, 2 }, { 65, 2 }, { 66, 2 }, { 68, 2 }, { 72, 2 },
            { 74, 2 }, { 77, 2 }, { 78, 2 }, { 79, 2 }, { 80, 2 }, { 81, 2 },
            { 82, 2 }, { 83, 2 }, { 85, 2 }, { 89, 2 }, { 91, 2 }, { 94, 2 },
            { 95, 2 }, { 96, 2 }, { 97, 2 }, { 98, 2 }, { 99, 2 },
            { 100, 2 }, { 102, 2 }, { 106, 2 }, { 108, 2 }, { 111, 2 },
            { 112, 2 }, { 113, 2 }, { 114, 2 }, { 115, 2 }, { 116, 2 },
            { 117, 2 }, { 119, 2 }, { 123, 2 }, { 125, 2 }, { 128, 2 },
            { 129, 2 }, { 130, 2 }, { 131, 2 }, { 132, 2 }, { 133, 2 },
            { 134, 2 }, { 136, 2 }, { 140, 2 }, { 142, 2 }, { 145, 2 },
            { 146, 2 }, { 149, 2 }, { 151, 2 }, { 153, 2 }, { 157, 2 },
            { 159, 2 }, { 162, 2 }, { 170, 2 }, { 176, 2 }, { 180, 2 },
            { 181, 2 }, { 182, 2 }, { 183, 2 }, { 184, 2 }, { 185, 2 },
            { 187, 2 }, { 191, 2 }, { 193, 2 }, { 194, 2 }, { 196, 2 },
            { 199, 2 }, { 205, 2 }, { 206, 2 }, { 207, 2 }, { 208, 2 },
            { 209, 2 }, { 210, 2 }, { 211, 2 }, { 217, 2 }, { 218, 2 },
            { 219, 2 }, { 229, 2 }, { 230, 2 }, { 231, 2 }, { 241, 2 },
            { 242, 2 }, { 243, 2 }, { 244, 2 }, { 245, 2 }, { 247, 2 },
            { 253, 2 }, { 266, 2 }, { 268, 2 }, { 277, 2 }, { 280, 2 },
            { 289, 2 }, { 291, 2 }, { 292, 2 }, { 293, 2 }, { 296, 2 },
            { 297, 2 }, { 298, 2 }, { 300, 2 }, { 306, 2 }, { 310, 2 },
            { 311, 2 }, { 312, 2 }, { 322, 2 }, { 323, 2 }, { 324, 2 },
            { 334, 2 }, { 336, 2 }, { 338, 2 }, { 339, 2 }, { 340, 2 },
            { 343, 2 }, { 349, 2 }, { 350, 2 }, { 351, 2 }, { 353, 2 },
            { 355, 2 }, { 364, 2 }, { 365, 2 }, { 366, 2 }, { 367, 2 },
            { 368, 2 }, { 371, 2 }, { 375, 2 }, { 377, 2 }, { 379, 2 },
            { 380, 2 }, { 381, 2 }, { 383, 2 }, { 387, 2 }, { 390, 2 },
            { 391, 2 }, { 392, 2 }, { 393, 2 }, { 395, 2 }, { 399, 2 },
            { 407, 2 }, { 412, 2 }, { 413, 2 }, { 414, 2 }, { 415, 2 },
            { 416, 2 }, { 417, 2 }, { 419, 2 }, { 423, 2 }, { 427, 2 },
            { 428, 2 }, { 431, 2 }, { 440, 2 }, { 442, 2 }, { 444, 2 },
            { 445, 2 }, { 446, 2 }, { 447, 2 }, { 448, 2 }, { 449, 2 },
            { 451, 2 }, { 455, 2 }, { 469, 2 }, { 470, 2 }, { 471, 2 },
            { 474, 2 }, { 476, 2 }, { 480, 2 }, { 481, 2 }, { 482, 2 },
            { 483, 2 }, { 484, 2 }, { 488, 6 },
            { 489, 6 }, { 490, 10 }, { 491, 2 }, { 493, 6 }, { 494, 6 },
            { 495, 6 }, { 496, 6 }, { 499, 6 }, { 508, 2 }, { 509, 6 }, { 512, 2 }, { 513, 2 }, { 514, 2 },
            { 515, 2 }, { 520, 2 }, { 521, 2 },
            { 522, 2 }, { 523, 2 }, { 524, 2 }, { 525, 2 }, { 526, 2 },
            { 527, 2 }, { 528, 2 }, { 529, 2 }, { 533, 2 }, { 534, 2 },
            { 535, 2 }, { 536, 2 }, { 538, 2 }, { 540, 4 }, { 541, 4 },
            { 542, 8 }, { 544, 2 }, { 545, 2 }, { 547, 2 }, { 548, 2 },
            { 549, 2 }, { 550, 2 }, { 551, 2 }, { 552, 2 }, { 555, 2 },
            { 556, 2 }, { 557, 2 }, { 558, 2 }, { 559, 2 }, { 560, 2 },
            { 576, 2 }, { 577, 2 }, { 578, 2 }, { 579, 2 }, { 580, 2 },
            { 581, 2 }, { 583, 2 }, { 587, 2 }, { 588, 2 }, { 598, 2 },
            { 599, 2 }, { 607, 4 }, { 609, 4 }, { 610, 2 }, { 611, 2 },
            { 613, 2 }, { 614, 2 }, { 616, 6 }, { 617, 6 }, { 618, 2 },
            { 691, 6 }, { 692, 6 }, { 694, 6 }, { 695, 6 }, { 702, 6 },
            { 703, 6 }, { 704, 6 }, { 705, 6 }, { 706, 6 }, { 709, 2 },
            { 712, 6 }, { 713, 6 }, { 714, 6 }, { 715, 6 }, { 716, 6 },
            { 717, 6 }, { 719, 6 }, { 720, 6 }, { 721, 6 }, { 722, 6 },
            { 739, 6 }, { 740, 6 }, { 752, 6 }, { 768, 6 }, { 771, 6 },
            { 772, 6 }, { 773, 6 }, { 774, 6 }, { 776, 6 }, { 779, 6 },
            { 780, 6 }, { 782, 6 }, { 783, 6 }, { 784, 6 }, { 785, 6 },
            { 786, 6 }, { 787, 6 }, { 789, 2 }, { 790, 2 }, { 791, 2 },
            { 792, 2 }, { 793, 2 }, { 797, 6 }, { 798, 2 }, { 799, 2 },
            { 800, 2 }, { 801, 2 }, { 802, 6 }, { 816, 6 }, { 817, 6 },
            { 818, 6 }, { 819, 6 }, { 820, 6 }, { 821, 6 }, { 822, 10 },
            { 823, 6 }, { 824, 6 }, { 827, 6 }, { 828, 6 }, { 829, 6 },
            { 830, 6 }, { 831, 6 }, { 832, 4 }, { 848, 6 }, { 849, 6 },
            { 850, 6 }, { 851, 6 }, { 852, 6 }, { 853, 6 }, { 855, 6 },
            { 856, 6 }, { 859, 6 }, { 860, 6 }, { 861, 6 }, { 862, 6 },
            { 863, 6 }, { 864, 4 }, { 865, 4 }, { 866, 6 }, { 868, 6 },
            { 869, 6 }, { 870, 6 }, { 871, 6 }, { 873, 6 }, { 874, 6 },
            { 875, 6 }, { 876, 6 }, { 877, 6 }, { 879, 6 }, { 881, 6 },
            { 882, 6 }, { 885, 6 }, { 896, 6 }, { 897, 6 }, { 898, 6 },
            { 899, 6 }, { 900, 6 }, { 901, 6 }, { 902, 6 }, { 903, 6 },
            { 904, 6 }, { 907, 6 }, { 908, 6 }, { 910, 6 }, { 911, 6 },
            { 912, 6 }, { 928, 6 }, { 929, 6 }, { 930, 6 }, { 931, 6 },
            { 932, 6 }, { 933, 6 }, { 934, 6 }, { 935, 6 }, { 936, 6 },
            { 939, 6 }, { 940, 6 }, { 941, 6 }, { 942, 6 }, { 943, 6 },
            { 944, 6 }, { 945, 6 }, { 946, 6 }, { 947, 6 }, { 948, 6 },
            { 950, 4 }, { 952, 2 }, { 953, 2 }, { 956, 6 }, { 957, 6 },
            { 958, 6 }, { 960, 6 }, { 961, 2 }, { 962, 2 }, { 976, 4 },
            { 977, 4 }, { 978, 4 }, { 979, 4 }, { 980, 4 }, { 981, 4 },
            { 982, 4 }, { 983, 4 }, { 984, 4 }, { 987, 4 }, { 988, 4 },
            { 989, 4 }, { 990, 4 }, { 992, 4 }, { 1008, 4 }, { 1009, 4 },
            { 1010, 4 }, { 1011, 4 }, { 1012, 4 }, { 1013, 4 }, { 1014, 4 },
            { 1015, 4 }, { 1016, 4 }, { 1019, 4 }, { 1020, 4 }, { 1021, 4 },
            { 1022, 4 }, { 1023, 4 }, { 1024, 4 }, { 1025, 4 }, { 1026, 4 },
            { 1027, 4 }, { 1028, 4 }, { 1029, 6 }, { 1030, 6 }, { 1031, 6 },
            { 1033, 6 }, { 1034, 6 }, { 1035, 6 }, { 1036, 6 }, { 1037, 4 },
            { 1056, 10 }, { 1057, 10 }, { 1059, 10 }, { 1060, 10 },
            { 1061, 10 }, { 1062, 10 }, { 1064, 10 }, { 1067, 10 },
            { 1068, 10 }, { 1072, 10 }, { 1088, 10 }, { 1091, 10 },
            { 1092, 10 }, { 1093, 10 }, { 1094, 10 }, { 1095, 10 },
            { 1099, 10 }, { 1100, 10 }, { 1102, 10 }, { 1103, 10 },
            { 1104, 10 }, { 1105, 10 }, { 1106, 10 }, { 1108, 10 },
            { 1109, 4 }, { 1110, 4 }, { 1115, 6 }, { 1117, 6 }, { 1118, 6 },
            { 1119, 6 }, { 1168, 6 }, { 1169, 6 }, { 1170, 6 },
            { 1171, 6 }, { 1172, 6 }, { 1173, 6 }, { 1175, 6 }, { 1176, 6 },
            { 1179, 6 }, { 1180, 6 }, { 1181, 6 }, { 1183, 6 }, { 1185, 4 }, { 1186, 4 }, { 1200, 6 }, { 1201, 6 }, { 1202, 6 },
            { 1203, 6 }, { 1204, 6 }, { 1205, 6 }, { 1207, 6 }, { 1208, 6 },
            { 1211, 6 }, { 1212, 6 }, { 1213, 6 }, { 1215, 6 }, { 1232, 6 }, { 1233, 6 }, { 1234, 6 }, { 1235, 6 }, { 1236, 6 },
            { 1237, 6 }, { 1239, 6 }, { 1240, 6 }, { 1243, 6 }, { 1244, 6 },
            { 1245, 6 }, { 1247, 6 }, { 1248, 4 }, { 1249, 4 }, { 1250, 4 },
            { 1264, 6 }, { 1267, 6 }, { 1268, 6 }, { 1269, 6 }, { 1272, 6 },
            { 1275, 6 }, { 1276, 6 }, { 1277, 6 }, { 1279, 6 }, { 1296, 6 }, { 1297, 6 }, { 1298, 6 }, { 1299, 6 }, { 1300, 6 },
            { 1301, 6 }, { 1304, 6 }, { 1307, 6 }, { 1308, 6 }, { 1313, 6 }, { 1314, 6 }, { 1344, 8 }, { 1345, 14 }, { 1378, 6 },
            { 1379, 6 }, { 1408, 4 }, { 1409, 4 }, { 1412, 4 },
            { 1413, 10 }, { 1414, 10 }, { 1416, 10 }, { 1417, 10 },
            { 1418, 10 }, { 1424, 10 }, { 1425, 10 }, { 1426, 10 },
            { 1427, 10 }, { 1428, 10 }, { 1429, 10 }, { 1430, 10 },
            { 1436, 10 }, { 1437, 10 }, { 1438, 10 }, { 1443, 10 }, { 1444, 10 }, { 1446, 10 }, { 1447, 10 },
            { 1448, 10 }, { 1449, 10 }, { 1454, 10 }, { 1455, 10 },
            { 1456, 10 }, { 1457, 10 }, { 1458, 10 }, { 1459, 10 },
            { 1460, 10 }, { 1461, 10 }, { 1466, 10 }, { 1467, 12 },
            { 1470, 12 }, { 1478, 10 }, { 1482, 2 }, { 1488, 6 },
            { 1489, 8 }, { 1490, 10 }, { 1491, 4 }, { 1492, 6 }, { 1493, 6 },
            { 1494, 6 }, { 1495, 6 }, { 1496, 8 }, { 1497, 10 },
            { 1498, 12 }, { 1499, 6 }, { 1500, 8 }, { 1501, 8 }, { 1510, 8 }, { 1511, 8 }, { 1512, 4 },
            { 1513, 4 }, { 1518, 4 },
            { 1519, 6 }, { 1521, 6 }, { 1522, 6 }, { 1523, 10 },
            { 1526, 10 }, { 1529, 10 }, { 1530, 10 }, { 1531, 14 },
            { 1532, 14 }, { 1533, 14 }, { 1535, 2 }, { 1536, 2 },
            { 1537, 2 }, { 1538, 12 }, { 1539, 12 }, { 1540, 2 },
            { 1543, 2 }, { 1552, 6 }, { 1553, 6 }, { 1554, 6 }, { 1555, 6 },
            { 1556, 6 }, { 1557, 6 }, { 1558, 6 }, { 1559, 6 }, { 1560, 6 },
            { 1563, 6 }, { 1564, 6 }, { 1565, 6 }, { 1566, 6 }, { 1567, 6 },
            { 1568, 6 }, { 1584, 6 }, { 1585, 6 }, { 1586, 6 }, { 1587, 6 },
            { 1588, 6 }, { 1589, 6 }, { 1590, 6 }, { 1591, 6 }, { 1592, 6 },
            { 1595, 6 }, { 1596, 6 }, { 1597, 6 }, { 1600, 6 }, { 1601, 6 }, { 1602, 6 }, { 1603, 6 }, { 1604, 6 },
            { 1605, 12 }, { 1610, 2 }, { 1611, 2 }, { 1612, 2 },
            { 1614, 2 }, { 1616, 2 }, { 1618, 2 }, { 1619, 2 }, { 1623, 6 }, { 1624, 2 }, { 1625, 2 }, { 1626, 6 },
            { 1627, 2 }, { 1628, 2 }, { 1629, 8 }, { 1631, 6 }, { 1632, 6 },
            { 1633, 6 }, { 1635, 2 }, { 1636, 2 }, { 1637, 2 }, { 1638, 2 },
            { 1646, 6 },
            { 1647, 6 }, { 1653, 2 }, { 1654, 2 }, { 1656, 2 }, { 1657, 2 },
            { 1658, 2 }, { 1660, 6 }, { 1661, 6 }, { 1662, 2 },
            { 1663, 2 }, { 1664, 6 }, { 1665, 10 }, { 1672, 4 },
            { 1673, 10 }, { 1681, 8 }, { 1682, 10 }, { 1683, 4 }, { 1684, 6 },
            { 1685, 6 }, { 1686, 2 }, { 1688, 2 }, { 1689, 2 }, { 1690, 2 },
            { 1691, 2 }, { 1693, 2 }, { 1698, 10 },
            // `CopyBytes`/`CopyBytesZero` (1037/1672, pinned above) repurpose RSI as the memcpy
            // pointer and restore it afterwards, so a handler walker reads nothing
            // and the corpus has never reached them. Their 2-byte count operand
            // comes from the published VB6 table, matched by symbol name and
            // confirmed in the handler: it restores rsi, fetches the next opcode
            // at +2 and advances 4.
            // The rows below were settled by operand structure in the runtime corpus.
            { 1113, 4 }, { 335, 6 }, { 582, 2 }, { 430, 2 }, { 1120, 2 }, { 1121, 6 }, { 596, 2 }, { 258, 2 },
            { 595, 2 }, { 67, 6 }, { 1431, 10 }, { 459, 2 }, { 606, 2 }, { 458, 2 }, { 358, 2 }, { 608, 2 }, { 1643, 3 }, { 257, 2 }, { 1439, 10 }, { 341, 2 }, { 1440, 10 }, { 472, 2 }, { 342, 2 }, { 473, 2 }, { 1469, 10 }, { 584, 2 },
            // 517/519 are DEAD, not unmeasured: the dispatcher enters every handler
            // with rax = the opcode index, the `3` twins overwrite it with a fetched
            // count before `add rsi, rax`, and the `4` twins do not -- they would add
            // 517 or 519 to the instruction pointer. The three bytes between are a
            // NOP. Nothing this compiler emits can be one of them.
            // 342/473 share a handler byte-for-byte with 341/472 (same RVA), so
            // their length follows by IDENTITY, not inference; 1469 jumps into the
            // same body as 1468, which the corpus confirmed at 10. 517/519
            // (GetRecOwner4/PutRecOwner4) stay EMPTY: their shared body adds a
            // register to RSI that nothing on the path loaded, and no record shape
            // tried -- fixed or dynamic strings, Variants, objects, Binary and
            // Random -- makes the compiler emit them. Empty rather than wrong.
            // The VCallBasic value forms (1232-1247) and VCallBasicCbFrame (1248) were
            // at 2, the fetch value. They are a CALL family, and their template is
            // exact: ThisVCallBasic (1168-1183) is entirely 6 with its CbFrame at 4,
            // and the Hresult members of both (498/499) are corpus-confirmed 6. Each
            // VCallBasic handler reads [rsi] and saves rsi identically to 498. So the
            // value forms are 6 and the CbFrame 4, mirroring the This-family. Corpus
            // has never reached them, so this is family+identity evidence, not a walk.
            // Through-call handlers the corpus has not reached: through-call rule, family-consistent.
            { 854, 6 }, { 867, 6 }, { 1112, 4 }, { 1419, 10 },
            // Adopted from the through-call rule and since WALKED PAST CLEANLY at an aligned boundary in
            // the runtime corpus -- the walk continued from the length to a real slot.
            { 486, 4 }, { 487, 4 }, { 506, 4 }, { 1411, 4 }, { 1474, 10 }, { 1475, 6 }, { 1476, 6 }, { 1607, 8 },
            { 1640, 4 }, { 1642, 4 }, { 1659, 4 }, { 7, 6 }, { 19, 6 }, { 31, 6 }, { 43, 6 }, { 55, 6 },
            { 73, 6 }, { 75, 2 }, { 76, 2 }, { 84, 6 }, { 90, 6 }, { 92, 2 }, { 93, 2 }, { 101, 6 },
            { 107, 6 }, { 109, 2 }, { 110, 2 }, { 118, 6 }, { 124, 6 }, { 126, 2 }, { 127, 2 }, { 135, 6 },
            { 141, 6 }, { 143, 2 }, { 144, 2 }, { 152, 6 }, { 158, 6 }, { 160, 2 }, { 161, 2 }, { 169, 6 },
            { 175, 6 }, { 186, 2 }, { 192, 2 }, { 212, 6 }, { 224, 6 }, { 236, 6 }, { 248, 6 }, { 260, 6 },
            { 272, 6 }, { 284, 6 }, { 290, 6 }, { 301, 6 }, { 313, 6 }, { 325, 6 }, { 331, 6 }, { 333, 6 },
            { 370, 2 }, { 378, 2 }, { 382, 2 }, { 394, 2 }, { 418, 2 }, { 450, 2 }, { 456, 2 }, { 457, 2 },
            { 460, 2 }, { 461, 2 }, { 463, 2 }, { 465, 2 }, { 466, 2 }, { 467, 2 }, { 468, 2 }, { 475, 2 },
            { 590, 2 }, { 612, 2 }, { 1376, 2 }, { 1377, 2 }, { 1613, 6 }, { 1615, 6 }, { 1617, 6 }, { 1644, 2 },
            { 1692, 2 },
            // THRULEN CANDIDATES: through-call and straight-line handlers the corpus
            // has not reached at an aligned boundary. Unverified: the handler read
            // has matched every byte-checked sibling, but nothing here has been read
            // from bytes. The offline solver promotes or rejects each as the corpus grows.
            { 1122, 8 }, { 1184, 4 }, { 1216, 4 }, { 1280, 12 }, { 1312, 12 }, { 1598, 4 }, { 1599, 4 }, { 1620, 4 },
            { 1621, 4 }, { 1639, 4 }, { 1641, 4 }, { 1666, 4 }, { 1667, 6 }, { 1668, 8 }, { 1669, 8 }, { 1670, 4 },
            { 1671, 4 }, { 1675, 10 }, { 1676, 8 }, { 1677, 8 }, { 1678, 8 }, { 1679, 10 }, { 177, 2 }, { 178, 2 },
            { 510, 2 }, { 511, 14 }, { 591, 2 }, { 592, 2 }, { 593, 2 }, { 594, 2 }, { 601, 2 }, { 602, 2 },
            { 603, 2 }, { 604, 2 }, { 605, 2 }, { 794, 2 }, { 795, 2 }, { 796, 2 }, { 1116, 6 }, { 1541, 2 },
            { 1542, 2 }, { 1608, 6 }, { 1674, 6 },
        };

        // LENGTHS THAT DEPEND ON THE OPERAND. `FFreeVar`/`FFreeStr`/`FFreeAd`
        // (1514-1516) release the Variant, String and object locals on exit:
        // opcode, a count word, then count/2 four-byte frame offsets -- 4 +
        // 2*count, in 87 of 87 corpus occurrences of `FFreeVar`. No
        // fixed pin could close a segment through it, which from outside looks
        // exactly like a fixed pin that never fits the hold-out.
        //
        // The named late calls (1502-1509) carry a list of argument-name ids and
        // the word at +2 is the byte count of what follows it: 4 + count, or
        // 6 + count for the two `LdVar` forms, whose 4-byte result slot the
        // count does not cover. `OnGoto`/`OnGosub` (1480/1481) skip a list of
        // four-byte branch targets the same way, and `GetRecOwner3`/`PutRecOwner3`
        // (516/518) a record payload -- all four were pinned at 2, which steps
        // into the middle of the list.
        //
        // FOUND BY A HANDLER SWEEP, not by the corpus: not one of those four
        // occurs in 3,004 traced procedures, so every instrument we own called
        // the table healthy. The signature is `add rsi, <register>` where the
        // register was fetched from `[rsi]`. `For`/`Next`/`Gosub`/`OnGoto` all
        // add a register to RSI too, but take it from elsewhere, because there it
        // is a BRANCH -- and `Resume` passes the naive test while REPLACING rsi
        // with a stored offset, so the fetch is what separates them.
        struct VarLength { std::uint16_t slot; std::uint8_t base; std::uint8_t unit; };
        constexpr VarLength kVarLength[] = {
            { 1514, 4, 2 }, { 1515, 4, 2 }, { 1516, 4, 2 },
            { 1502, 4, 1 }, { 1503, 6, 1 }, { 1504, 4, 1 },
            { 1505, 4, 1 }, { 1506, 4, 1 }, { 1507, 6, 1 },
            { 1508, 4, 1 }, { 1509, 4, 1 },
            { 1480, 4, 1 }, { 1481, 4, 1 },
            { 516,  4, 1 }, { 518,  4, 1 },
        };

        // WHICH SLOTS THE RUNTIME CORPUS HAS WALKED PAST, one bit per slot.
        // Generated offline: a set bit is a slot seen at an aligned boundary in a
        // procedure that parsed on to its exit, so its length has been exercised by
        // real compiler output. Everything pinned and NOT here rests on the handler
        // alone, so the walk counts a step through one and says so at disarm.
        // WALKED IS NOT CONFIRMED: a slot the corpus stepped over can still be
        // pinned wrong -- the walk only proves the procedures containing it
        // happened to close.
        constexpr std::uint8_t kCorpusWalked[] = {
            0x84, 0x00, 0x08, 0x8C, 0xC0, 0x08, 0x8C, 0x80, 0x18, 0x1A, 0x12, 0x34,
            0x28, 0x68, 0x44, 0xD0, 0x80, 0xA0, 0x79, 0x61, 0x07, 0x86, 0x00, 0x04,
            0xFD, 0x21, 0x17, 0x10, 0x81, 0x11, 0x70, 0x01, 0x16, 0x08, 0x81, 0x1F,
            0x4E, 0x3F, 0x44, 0x02, 0x24, 0xB8, 0x21, 0x00, 0x7E, 0x28, 0x05, 0x46,
            0xD0, 0x04, 0x7B, 0x68, 0x04, 0x67, 0x38, 0x2E, 0x04, 0xBF, 0x1E, 0x09,
            0xEA, 0x78, 0x1F, 0x05, 0x00, 0x00, 0x00, 0x5B, 0xC4, 0xE0, 0x00, 0x00,
            0x44, 0x61, 0x38, 0xC1, 0xD3, 0xE8, 0xDF, 0x0C, 0x00, 0x00, 0xFF, 0xF9,
            0x01, 0x00, 0xE7, 0xF9, 0xFC, 0x60, 0x02, 0x00, 0xF7, 0xF9, 0x00, 0x00,
            0xE7, 0x39, 0x18, 0x00, 0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xA0, 0xFC, 0x26, 0x00, 0x04, 0x61, 0x00, 0x00, 0x04, 0x21, 0x90, 0xF1,
            0x07, 0x00, 0x84, 0x81, 0x00, 0x00, 0x04, 0x20, 0x04, 0x01, 0x00, 0x00,
            0xE7, 0xF9, 0x00, 0x00, 0xE7, 0x79, 0xFC, 0x6E, 0x03, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x06, 0x00, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x86, 0xB1, 0x00, 0x00, 0x84, 0xA1, 0x00, 0x00, 0x00, 0x00,
            0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x99, 0x0B, 0xB8, 0xE0,
            0xE7, 0x02, 0x2E, 0xD8, 0x3F, 0x03, 0x0F, 0xCF, 0x41, 0x3F, 0xB1, 0x73,
            0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0xE0, 0xE2, 0x8A, 0xE0,
            0x85, 0x97, 0xFF, 0xCE, 0x01, 0x00, 0x8E, 0x14, 0x00,
        };
    }
}
