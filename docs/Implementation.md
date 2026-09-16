# Implementation — XRayXL

*The technical spec for the code that exists.*

> **This document describes what is built.** Where a mechanism is specified here
> it is in `src/` and under test. Design work that has not been built is
> not in this file — a spec that mixes the two stops being usable as a
> description of the program.

---

## Scope

> **When Excel's calculation engine calls a user-defined function — an XLL
> function or a VBA function — record one row saying which function ran, in
> which cell, on which thread, with what arguments, returning what, and for how
> long.**

| Part | | |
|---|---|---|
| **1. The vehicle** | The bare add-in and its command surface | Small, and entirely documented API |
| **2. Derivation** | What is derived, and what is not | One derived structure, on the VBA side only |
| **3. XLL UDF tracing** | The XLL tracer | **Zero derived addresses** — documented API throughout |
| **4. VBA UDF tracing** | The VBA tracer | A dispatch-table patch and a per-thread shadow stack |
| **5. Safety** | Applies from the first hook | Non-negotiable |

Anything not in those five parts is not in scope. **Deliberately left out:**
COM/RTD tracing; native built-in tracing (removed after it was measured not to
work); VBA the calculation engine did not call *as a target*, though macros,
events and form code are traced anyway because the interpreter hook is global;
a span model, parent IDs or causality trees; Chrome Trace Event or Perfetto
output; UDP transport, an out-of-process
collector, a watchdog, multi-session demux; presets and live session-time
reconfiguration; a hosted CLR or any managed code in the Excel process; 32-bit
Excel. Several are good ideas; they come back one at a time, as separate
decisions.

## Where the code is

| Part | Files |
|---|---|
| **1. The vehicle** | `src/dllmain.cpp`, `src/XRayXL.def`; `src/app/` — `session` (arm and disarm both sources), `commands` (the `XRayXL_*` exports and their guards), `traceparam` (Set/Get), `summary` (the trace summary), `paramparse`/`xlgrid` (their arguments and answers); `src/core/` — logging, the crash log, the Excel API wrappers, trace modes, and what both tracers share: `caller` (the calling cell), `render`, `text`, `clock`, `safemem` |
| **2. Derivation** | `src/vba/vbaderive.*` — locating and verifying the dispatch table |
| **3. XLL tracing** | `src/xll/` — `xllregistry` (enumerate), `xlltypeplan` (parse the registration), `xllhook` + `xllthunk.asm` (patch and wrap), `xlldecode` (read the values), `xlltrace` (build the row), `xllarm` (hook and unhook this source), `xllregister_watch` (add-ins loaded after arming) |
| **4. VBA tracing** | `src/vba/` — `vbapatch` + `vbathunk.asm` (swap the slots), `vbapcode` + `vbapcode_tables.h` (the walk, and the pinned lengths it walks by), `vbatrace` (the shadow stack) + `vbareport` (its disarm report), `vbaargs`/`vbaretdecode` (read the values), `vbaidentity`/`vbaproctable`/`vbatrailer` (name the procedure) |
| **5. Safety** | `src/core/crashlog.*`, `src/core/safemem.h`, and the SEH guards at every hook body |
| The output | `src/emit/` — `ring` (the buffer), `rowcsv` (formatting), `csv` (the writer) |
| Vendored | `src/third_party/` — MinHook, and Microsoft's `xlcall.h` |

**`src/xll/` and `src/vba/` do not include each other.** What they share lives in `src/core/`; what drives both lives in `src/app/`. The `xll` folder is deliberately absent from the include path, so a cross-folder include has to be written out and is easy to grep for.

**Where to start reading.** Part 3 first: the XLL tracer derives nothing, so it
shows the shape of the whole design — hook, decode, emit — with none of the
machinery. Then Part 4, which answers the same questions from a much harder
place. Part 2 makes more sense once you have seen why the VBA side needs it,
and Part 5 applies throughout.

## The core problem, precisely

The thing a calculation diagnostic most needs — **which cell caused this to
run** — is not a property of the function. It lives in the calculation engine.

For XLL functions that is a solved problem with a documented answer. For VBA it
is not present in the interpreter's own state in any form, which is why the two
sides are sized so differently despite producing the same row.

So the design's problem is: *find a vantage point per technology from which
identity, arguments, result, timing and the triggering cell are all
simultaneously available — and find it without writing down a single
build-specific address.*



# Part 1 — The vehicle

## Delivery: our own XLL, nothing else

Ships as a normal XLL, installed the way any XLL is (`XLSTART` or Trust Center).
**No separate injector, no `CreateRemoteThread`, no external DLL injection.**
Being loaded as an XLL already puts the code inside Excel's address space, which
is all that any later hook needs. This keeps the install story simple and
meaningfully lowers the AV/EDR suspicion profile against a standalone injector.



## The command surface



Arming is a registered command:

| Call | Does |
|---|---|
| `Application.Run("XRayXL_Arm")` | Derive, verify and install hooks; fail closed if verification does not pass. |
| `Application.Run("XRayXL_Disarm")` | Disable the hooks, restoring the original code; returns the number of rows dropped. |

What gets traced is set, and the session inspected, through registered functions:

| Function | Does |
|---|---|
| `XRayXL_SetTraceParam(Source, Name, Value)` | Set one setting. `Source` is `XLL` or `VBA`, and an omitted `Source` addresses both; `Name` is `DEPTH` (`OFF`/`TOP`/`ALL`), `ARGS`, `RETVAL` or `OBJECTS`. |
| `XRayXL_GetTraceParam(Source, Name)` | Read one setting back, or a grid of them when an argument is omitted. Volatile. |
| `XRayXL_GetTraceSummary(Filter)` | List what was traced, with live call counts. Volatile. |
| `XRayXL_IsArmed()` | `TRUE` while either source is armed. |

`XRayXL_FaultProbe`, which faults on purpose to test the crash handler, is
registered only when `XRAYXL_DIAG=1`.

Three settings take no `Source` — `BUFFERSIZE` (ring size in MB), `BUFFERWHENFULL`
(`DROP`/`PAUSE`) and `LOGLEVEL` (`DEBUG`/`INFO`/`WARNING`/`ERROR`). Every
per-source setting defaults to on, with `DEPTH=ALL` for both sources. **The
setters refuse while armed and refuse when the caller is a
cell** — a trace setting changed mid-calculation, or written by a formula, would
no longer describe the run it is attached to. (`LOGLEVEL` is the exception to
"refuse while armed": it touches only logging, so it is settable at any time.)

## Lifecycle

| Export | Does |
|---|---|
| `xlAutoOpen` | Register the commands and trace-parameter functions. Open the log. |
| `xlAutoClose` | Close the log. |
| `xlAutoFree12` | Free `XLOPER12`s we allocated. |
| `xlAddInManagerInfo12` | Name in the add-in manager. |

**That is the whole export surface** — the four lifecycle exports, plus the
registered commands and functions above.

## Export-surface discipline

Diagnostic exports accumulate: each one is added to answer a question, and
almost none is removed when the question is answered. An add-in that exports
150 entry points has stopped being the shape it is supposed to be.

**The rule: a diagnostic export is added deliberately, named for its question,
and deleted when the question is answered.** The export count is
reviewed regularly. It is a good proxy for whether the codebase is still the
shape it is supposed to be.

## Logging

`crashlog` and the `Log` (`core/log.{h,cpp}`) survive the strip. Being able to
see what the add-in did — and, in the field, what it did before it stopped — is the
one facility you never want to be building under pressure. They are small and
they are kept.

**They are two different writers, not two files doing one job**, and the split is
forced rather than stylistic. `crashlog` runs inside an unhandled-exception
filter, where the heap may already be corrupt — that may be *why* it is running —
and where a lock could be held by the thread that just died. So it allocates
nothing, takes no lock, uses no CRT formatting and writes its hex by hand through
raw `CreateFileW`/`WriteFile`, opening and closing per append so the bytes are on
disk before the next instruction. The `Log` is ordinary application logging:
`std::ofstream`, a mutex, real formatting.

**The `Log` is levelled.** Four levels — DEBUG, INFO, WARNING, ERROR — in a
Log4Net line format (`2026-09-07 10:53:07,123 [ 4812] INFO - message`). The
default is INFO. The startup threshold is overridable by the `XRAYXL_LOGLEVEL`
env var, **read once** at load; a runtime `LOGLEVEL` trace-param (Get/Set) moves
it thereafter and, unlike the trace setters, is **settable while armed**.

Developer diagnostics in the disarm report are gated behind the `XRAYXL_DIAG` env
var (default off) — a separate switch from `LOGLEVEL`.

Their truncation semantics are opposite on purpose too. The `Log` truncates
on open, because it is the narrative of *this* session. A crash log that
truncated on open would erase the record of the crash that killed the last one —
which is the record you wanted. **Do not merge them**: the merge would mean
either the `Log` losing its formatting, or the crash log gaining allocation,
and the second is the bug.

**The vehicle's baseline:** Excel starts and loads the XLL, the commands and
trace-parameter functions register and can be called through `Application.Run`,
and Excel closes cleanly. Nothing is hooked, nothing is derived, nothing is
patched, and no file is written but the log.

---

# Part 2 — What is derived, and what is not

**Two of the three mechanisms derive nothing at all.** Saying which is which is
the first thing to know about this design, because the cost, the risk and the
failure modes all follow from it.

## The XLL side derives nothing

Excel's registration table names every registered XLL function and resolves each
to its own address, and that is a documented API. The tracer enumerates the
table, parses each registration string into a decode plan, and hooks the
function it was handed.

There is no image scan, no stack walk, no candidate scoring and no calibration
pass. **No address in the XLL tracer comes from anywhere but an API answer.**
Everything in Part 3 rests on that. The one hook on `EXCEL.EXE` itself, the
registration watch on `MdCallBack12`, takes its address from `GetProcAddress`
like everything else.

## The VBA side derives one thing, by shape

The p-code opcode dispatch table in `VBE7.DLL` is the single derived structure in
the product, and it is found by **counting, not by searching for bytes**: scan
the module for runs of qwords that point back inside it, keep runs longer than
255 entries, and accept the candidate whose **two beginning-of-statement slots
hold the same handler**. That check alone disambiguated correctly on every build
tested, costs two loads, and needs no symbol.

> **Recognising a structure by its population and shape is not a byte
> signature.** A scanner hunting an instruction sequence would be a signature
> database wearing a different hat, and this project does not have one. What is
> recognised here is a table's *shape* — how many entries, how many distinct
> handlers, which slots agree — which is why it survived a fourteen-year gap
> between builds where the table's own address had moved by megabytes.

**The scan is guarded.** Locating the table dereferences every qword of VBE7's
image — about ten megabytes, the largest read in the subsystem. A discardable
section or an uncommitted tail would fault out of the arm handler, so reads go
through the `Image` abstraction (`Read`/`Scan`/`Base`/`SizeOfImage`) rather than
a raw pointer; on a fault it skips the rest of that page and counts
`ImageUnreadable` rather than taking an exception per qword.

## What derivation produces

`vba::SlotSet` (`vba/vbaderive.h`) is the value the rest of the VBA tracer consumes:

| Field | |
|---|---|
| `found` | A table was located at all |
| `verified` | Every structural check passed. **Arming refuses without it** |
| `PatchSite[]` | Per slot: the slot index, the `handlerRva` it currently holds, its `role` (`bos`, `exit`, `raise`, `end`) and its exit repeat-group |

**Sites carry RVAs, never absolute addresses.** ASLR moves the load base on every
start, so a site records an offset into the named module and the base is added at
runtime.

**A slot index is not an address**, which is why pinning one is not a hardcoded
offset. Across 41 VBE7 binaries, 2012 to 2026, the two beginning-of-statement
slots hold the same handler on 41/41, and all 25 exit slots hold one handler per
repeat-group on 41/41. The table *is* the interface.

## Declines are an enumeration, not a boolean

`vba::Decline` names why a derivation produced nothing — module not loaded, no
table found, wrong slot count, a non-uniform repeat-group, an unreadable image.
An instrument that cannot distinguish "never looked" from "looked and found
nothing" reports the two as the same sentence, and that has cost real days.

## Fail closed

Arming refuses if the table is not 1700 slots, or if any exit repeat-group is not
uniform, or if `verified` is not set. **A disarmed session is a normal outcome
and must never be silent.**

**Finding a table is not the same as being able to read it.** The p-code length
table is pinned rather than derived, so three separate gates stand between the
two — and each degrades a different thing rather than failing the whole arm:

| Gate | Says | On mismatch |
|---|---|---|
| `kPartitionHash` (`vba/vbaderive.cpp`) | The slot indices still mean what the table thinks: for every slot, the lowest slot sharing its handler, hashed | Types degrade off, values keep flowing |
| The clean-walk fraction (`g_cleanWalks`, `vba/vbapcode.cpp`) | The pinned instruction lengths actually work on this build — a walk that stays synchronised to a procedure's exit | Reported; suspect lengths are named by opcode |
| `kCorpusWalked` (`vba/vbapcode.cpp`) | Which slot lengths running code has ever exercised, one bit per slot | A length no corpus has walked past is flagged as unconfirmed |

The distinction matters: a derivation that finds an address and a decoder that
reads it correctly are different claims, and only the second is worth trusting a
trace to.


# Part 3 — XLL UDF tracing

**Goal: when Excel's calculation engine calls a function registered by any XLL
add-in, an entry record and an exit record come out.** The entry record carries
the decoded arguments; the exit record carries the decoded return value.

## Where to stand: at the function, not at the call site

The obvious place to intercept an add-in call is **Excel's internal call
site** — the one routine every add-in call passes through. It is undocumented,
so finding it means stack-walking out of probe functions, and everything
expensive follows from that: `.pdata` unwinding, candidate scoring, a
calibration pass, and the **stack-walk hazard** — published work on 64-bit
Office reports it registering a dynamic function table whose callback calls
`TerminateProcess`, so a walk that resolves an address inside that region kills
Excel silently, with nothing to attribute it to.

**So the tracer does not stand there.** The registration table names every registered
function and resolves each to its own address, and that is a documented API.
So we hook **the registered function itself**.

| | Call-site hook (old) | Function hook (now) |
|---|---|---|
| Derived addresses | 1, from stack walks | **0** |
| Stack walking | every derivation | **none** — the stack-walk hazard does not arise |
| Identity | runtime address→registration lookup | **free** — we know what we hooked |
| Signature | inferred | exact, from the registration string |
| Blast radius of a bug | every add-in call in the process | one function |

The cost, stated honestly: **N patches rather than one**, in third-party
modules; and functions registered *after* we arm need picking up.

> **Coverage is bounded by enumeration, and that must be counted.** A call-site
> hook sees every add-in call whether or not we knew about the function. A
> function hook sees exactly what we enumerated and armed. Every function we
> could not resolve, could not hook, or hooked after it had already been called
> is a **decline**, and declines are counted and reported — otherwise "no rows"
> and "we never looked" are the same sentence.

## One mechanism, for every XLL

**Hook the address Excel calls, with an inline prologue detour. That is the
whole rule.** There is no classification step, no per-framework path, and
nothing in the tracer that can tell one add-in's build from another's.

Excel resolves an export and calls the address. It cannot see what shape the
code there is, and neither do we: MinHook deals with whatever instructions it
finds, whether that is a real prologue, a linker's `E9 rel32` thunk, or a
6-byte `FF 25` indirect jump into a data slot.

> **There is deliberately no per-shape branch**, and the temptation is real:
> classify the export by its first bytes, and hook a jump table's data slot on
> the reasoning that a 6-byte thunk cannot be detoured safely. Two measurements
> say don't. Ordinary C functions are exported as `E9` thunks too — incremental
> linking does that to everything, so a "jump table export" is not a property of
> any framework — and hooked the same way **both shapes work**: 37 suite cases
> across both, one mechanism, zero declines. A branch here would buy nothing and
> make the tracer look as though it knew something about particular add-ins. If a
> binary shape ever does defeat the detour, the special case comes back with the
> measurement beside it.

**A prologue detour, never a call-site patch.** `EXCEL.EXE` runs with Control
Flow Guard over 58,474 registered call targets. Patching a call site would be
rejected and the process terminated. A prologue detour is never seen by CFG,
because Excel's indirect call still targets the original registered address and
only then reaches our patched bytes.

## Functions registered after arming

An add-in that registers a function after `XRayXL_Arm` — loaded later, or
registering lazily — would be invisible to an enumeration taken at arm time. So
arming also detours `MdCallBack12`, the `EXCEL.EXE` export every add-in's
`Excel12` call reaches. The detour notes `xlfRegister` calls and hands them to a
worker thread, which hooks each new function the same way (`xllregister_watch.cpp`,
`ArmLate` in `xllarm.cpp`). A registration Excel refused is not hooked, and
disarm disables the watch before anything else. `XRAYXL_NOREGWATCH=1` before
launching Excel turns the watch off.

## Entry and exit as two records

One call produces **two records** sharing a span id.

**Why two rather than one row with a duration.** An entry with no matching exit
is exactly what a hang or a crash looks like, and it is information the product
exists to surface. A single row emitted at exit cannot represent a call that
never finished.

| | carries |
|---|---|
| **entry** | span id, thread id, entry timestamp (`QueryPerformanceCounter`), function identity, calling cell + sheet, **decoded arguments** |
| **exit** | span id, exit timestamp, **decoded return value**, how it ended |

Span id is `(thread id, per-thread counter)` — composed in TLS, so there is no
cross-thread atomic and no contention on the hot path.

## Identity — and why the export name is not the answer

The registration table's procedure column is `pxProcedure`, **the export name**.
For Excel-DNA that is `f0`, `f1`, `f2`… — measured, not inferred. Anyone
reading `f37` learns nothing, and learns it confidently, which is the failure
mode this project keeps meeting.

The real name is `pxFunctionText`, and there is a documented route to it inside
the C API: `xlfRegister` defines a hidden name that evaluates to the
registration id, and `xlfGetDef` maps a definition back to a name with
`pxTypeNum = 2` for hidden names only.

**Resolve the name once, at arm time, and store it.** The hot path must never
do this.

Excel-DNA also registers two infrastructure functions of its own
(`RegistrationInfo`, `SyncMacro`). The tracer does not single them out: they are
traced like any other registration.

## The registration string — parse it once, correctly

The type string decides both what the arguments are and where they sit. Parsing
it is the single most defect-prone part of the XLL tracer, and the defects are
**measured**, not hypothetical.

**Consume a trailing `%` as part of the code.** The Excel-2007 types are two
characters — `C%`, `D%`, `F%`, `G%`, `K%`, `O%`.

**Track the ABI slot index separately from the type index.** `O`/`O%` is one
type code occupying **three** ABI slots (`rows`, `columns`, `array`).

Measured against a real Excel-DNA add-in — `double f(string, double[], double)`
registers as **`QD%K%E`**, which a per-character walk reads as **five arguments
instead of three**. Note that `QQ` and `QEE` come out right *by accident*,
because they contain no `%`: the defect bites every signature carrying a string
or an array, and no other.

**Where the argument sits — position picks the index, type picks the register
file.** `ireg[position]` for a pointer or integer type, `xmm[position]` for `B`.
These are not competing numbering schemes: Microsoft's own example puts the
second argument of `f(int, double, …)` in `XMM1`, not `XMM0`. Beyond the fourth,
`rspAtEntry + 0x28 + (n-4)*8`.

**Decode per code, from the measured table of type codes.** The traps worth
restating here:

- **Four string conventions**, not one — null-terminated vs counted, ASCII vs
  UTF-16. A single reader is right for two of six, and **the range guard does
  not catch the other four**: two ASCII characters read as one `WCHAR` land
  inside the permitted range, so the reader copies that many wide characters out
  of a buffer that never held them. Fabricated, readable, complete-looking text.
- **`FP` (`K`) and `FP12` (`K%`) differ only in the width of `rows`/`columns`,**
  and the payload starts at +8 in both — so a cross-decode does not crash, it
  reports **zero cells** with an absurd row count beside it.
- **`X` marks the function asynchronous.** See below.

## The return value — switch on `typeText[0]`

The registered return code is the first character of the type string, and it
must actually be **read**. Assuming every return is an `LPXLOPER12` in `rax` is
wrong for most of the table below.

| registered | where the answer is |
|---|---|
| `Q` `U` | `RAX` → `XLOPER12` — 32 bytes, `xltype` at **+24** |
| `P` `R` | `RAX` → **`XLOPER`** — 24 bytes, `xltype` at **+16**. A different struct. |
| `B` | **`XMM0`** |
| `A` `I` `J` `H` | `RAX` as an integer |
| `C` `D` `C%` `D%` `K` `K%` `E` `L` `M` `N` | `RAX` as a **raw pointer** |
| `F` `G` `F%` `G%`, digits `1`–`9` | modify-in-place — the answer is in an **argument** |
| `X` | **there is no return value yet** |

**The pointer row is the dangerous one and the guard makes it worse.** A
plausibility check that rejects small integers *passes* a genuine `wchar_t*`, so
the bytes 24 into somebody's string buffer get used as an `xltype` — a
complete-looking row carrying a fabricated value.

**Async is not a value problem, it is a timing lie.** An `X` in the type string
means the function is declared `void`, returns nothing at the call, and delivers
later through `xlAsyncReturn`. Timing the dispatch reports a four-second call as
microseconds — for a tool whose purpose is finding why a spreadsheet is slow,
that is worse than not tracing it. **Mark the span async and emit no duration**,
closing it only if and when the result arrives.

## Getting the exit edge without rewriting a return address

The cheap way to take an exit is to **overwrite the callee's return address**.
That is the origin of the **shadow-stack hazard**: a CET mismatch is a fast-fail
that bypasses SEH, and neither the guards nor the circuit breaker can see it.

**The tracer wraps instead.** Our thunk performs a genuine `call` into the original
target and regains control on its return. The shadow stack records our push and
matches it on `ret` — a legitimate call/ret pair, so **CET is satisfied by
construction** rather than by an exemption we must not disturb.

The address to call back into is MinHook's trampoline.

Forwarding arguments needs the arity, and **we have it from the registration
string before we arm**. Arity ≤ 4 needs nothing but shadow space; beyond that,
copy `n-4` stack qwords. One shared assembly thunk plus a per-function context
covers both, with the context reached through a volatile, non-argument register
(`r10`/`r11`).

> **`XRayXL64.xll` must still not be built `/CETCOMPAT`** — the exit thunk is
> no longer the reason, but the rule is cheap and the exemption is load-bearing
> for anything that later does rewrite a return address. See Part 5.

## The calling cell

`xlfCaller`, asked from inside the hook where Excel believes it has called the
add-in, returns the calling cell — **including for a thread-safe function under
multithreaded calculation**, which the documentation carves out explicitly.

It returns an `xltypeSRef`, which carries **no sheet**; `xlSheetNm` on that same
reference supplies the sheet as a *name*, which is what the record wants.

**Stay inside the documented-safe subset:** read row and column off the
reference, then call `xlSheetNm`/`xlSheetId`. **Never `xlCoerce` it** — that is
the one named hazard, failing with `xlretUncalced` on an uncalculated cell.
`xlfCaller`'s result **must be freed with `xlFree`**, and so must the sheet name.

**Where the cell cannot be established, report none.** Excel calls functions for
its own purposes with no calling cell, and those frames read as row 0, column 0
— which, taken at face value, reports a confident "A1".

## Multithreaded calculation

Every consequence here is a from-day-one requirement, not a later optimisation:

- **Every hook body is reentrant.** Multiple threads can be inside the same
  hooked function simultaneously. All state is **thread-local**; a shared
  counter or buffer would corrupt spans across threads.
- An XLL function runs on a worker thread only if registered thread-safe (`$` in
  its type text). The flag is parsed and skipped; the declared thread-safety is
  **not** recorded.
- **Every record carries a thread id**, which shows what ran where.
- The multithreaded-calculation configuration is **not** recorded.

## The hot path is sacred

Inside the hook, on the traced thread:

- **Nothing allocates, locks or blocks *per traced call*, except where a setting
  asks for it.** One-off per-thread setup on first use is permitted: the 64 KB
  argument render buffer (`vba/vbatrace.cpp`) and the row scratch
  (`emit/csv.cpp`), each taken on a thread's first use, reused after, and never
  freed on a path that might be inside a hook. The settings that do cost a call
  are stated: the calling cell is resolved on every call (`xlfCaller`, `xlSheetNm`)
  and is not a setting -- the VBA tracer needs it to attribute an error that escapes
  into a cell; `OBJECTS` calls the object model (`Address`, `Count`, `Value2`, `Name`) for an
  object argument or result; `BUFFERWHENFULL=PAUSE`, the default, makes the
  traced thread wait when the ring fills, until it is half empty; and
  `BUFFERSIZE=0` writes each row under a lock.
- **Per-thread state via TLS only.**
- **SEH-wrapped.** The VBA hooks also have a circuit breaker; an XLL recorder
  fault loses that row, is counted, and is reported at disarm.
- `QueryPerformanceCounter` for timing, never `GetTickCount`.
- **Everything derivable at arm time is derived at arm time** — the parsed type
  plan and the resolved name. The hot path reads a struct.

## Arming, and failing closed

1. Enumerate registrations; resolve each to an export address.
2. Parse the type string into a decode plan; resolve the real name.
3. Install; count and report everything declined.

**Disarm disables, and never removes.** A detour is created once per export and
kept, disabled, when the session ends: a thread still inside it, on a thread of
the add-in's own rather than Excel's, calls through its trampoline, so the
trampoline cannot be freed. Arming the export again enables the same detour.

**No other DLL is held in memory.** XRayXL pins only itself. Before any write into
another module, a detour or VBE7's dispatch table, it checks that the image it
patched is still at that address: the same base, link timestamp and size, and for
a detour the export's first bytes. An add-in unloaded while armed is left
untouched at disarm, and one loaded again unpatched at the same address is patched
again.

The decoders are verified by the test suites rather than at arm time: their
workbooks and `TracedAddin` plant arguments, cells and return values whose
answers are known. A disarmed session is a normal outcome and must never be silent.

## Getting the records out

Rows are built on the traced thread and deposited, as bytes, into a bounded
lock-free ring (`emit/ring.cpp`; `BUFFERSIZE`, default 64 MB). A writer thread
drains it to the trace file, so file I/O stays off the calculating thread. A
full ring either makes the producers wait until it is half empty
(`BUFFERWHENFULL=PAUSE`, the default) or drops the row (`DROP`); a row larger
than the whole ring is always dropped.
Drops are counted and reported by `XRayXL_Disarm` and the disarm log line, never
as rows. `BUFFERSIZE=0` writes each row synchronously instead, which is what a
crash hunt wants: everything up to the crash is on disk.

## Testing the XLL tracer

The principles that decide whether a suite is worth running — triggers and
formula state as dimensions, `AsLoaded`, a degraded rung being fatal for a test
run, and both add-in shapes — are in [Testing.md](./Testing.md), with the
XLL-specific type codes that bite. They are not repeated here.

## What arming costs

**Measured** at every arm and written to the log. From a sweep session with 52
registered functions:

| Side | Cost of one arm |
|---|---|
| VBA | **~30 ms** — derive and verify the table, pin the p-code lengths, patch 29 slots |
| XLL | **~10 ms** for 46 hooked functions |

```
arm cost: total 10ms = enumerate 1ms + name resolution 2ms
          (xlfRegisterId 0ms + xlfGetDef 1ms over 46 calls, 46 resolved)
          + hook install 6ms
```

The split is logged because a total cannot say which part is expensive. Hook
installation dominates, and the reason is in MinHook's own source: enabling one
target suspends every thread in the process, patches, and resumes. So every hook
is queued and applied with one `MH_ApplyQueued` — one freeze for the whole set —
and removal is batched the same way. The friendly name resolves for every
function (46 of 46 above).


---

# Part 4 — VBA UDF tracing

## The design in one paragraph

**Patch the p-code dispatch table — not code — to learn which procedure is
running and when it starts and stops. Keep depth, nesting and recursion in a
per-thread shadow stack. Ask Excel for the calling cell, and take the return
type from the exit opcode.** No single interception point answers the whole
question, and trying to make one do so is what stalled this for a long while.

| Mechanism | Answers | Cost / risk |
|---|---|---|
| **Dispatch-table patch** — BoS, 25 exit slots, raise and `End` | Which procedure is running; entry and exit edges | Fires per *statement*; patches a process-wide table |
| **TLS shadow stack** — keyed on the p-code trailer | Call tree, nesting depth, recursion | O(1) per event, no allocation |
| **`xlfCaller`, from inside the hook** | The calling cell | A call into Excel, so taken per activation and never per statement |

## Why patch the table and not the code

The interpreter reaches every opcode handler through one table of function
pointers. Making that table writable and swapping entries for our own thunks
rewrites **nothing in the instruction stream**.

That removes, in one move, all four hazards this document otherwise carries for
the VBA side: Control Flow Guard, CET shadow stacks, unwind metadata
invalidated by a displaced prologue, and the blast radius of detouring a function
every macro in the process runs through. **Disarming is a pointer write back** —
exact, and verifiable against the originals we recorded.

This supersedes the earlier idea of *wrapping "the interpreter function"* to
bracket VBA execution. That presumes there is exactly one function to wrap, which
is measured as **an era-specific property** — true since 2020, false for most of
VBE7's life.

## The slots are a legitimate constant

Normally a hardcoded offset is a bug by this project's rules. Slot indices are
the exception, and it is measured rather than argued: across **41 VBE7 binaries,
2012 to 2026**, the two beginning-of-statement slots hold the same handler
(`lblBEX_LargeBos`) on 41/41, and all 25 exit slots hold one handler per
repeat-group on 41/41.

The table **is** the interface. A slot index is not an address.

## The mechanism, step by step

1. **Arm.** Scan `VBE7.DLL` for runs of qwords pointing back inside the module;
   keep runs over 255. Accept the candidate whose **two BoS slots hold the same
   handler** — that check alone disambiguated correctly on every build tested,
   costs two loads, and needs no symbol. **Refuse to arm** if the table is not
   1700 slots, or if any exit repeat-group is not uniform.
2. **Patch.** `VirtualProtect` the table; replace the BoS handler and the 25 exit
   handlers with thunks that preserve volatile state, call our handler with
   `rsp`, and tail into the original. **Record every original pointer.**
3. **Identify.** At `rsp + 0xB8` sits the p-code trailer (`RTMI`) identifying the
   running procedure. Compare against this thread's current trailer; if unchanged,
   return. **This is the hot path, and it is a load and a compare.**
4. **Push or pop.** On a change, push a TLS frame carrying the trailer, a span id
   and a QPC stamp. An exit slot pops and emits the span. **Depth and recursion
   come from our stack, not the interpreter** — a recursive call has the same
   trailer, so nothing else can distinguish it.
5. **Resolve off the hot path.** The trailer chains to project, module and
   function name, ending in a linear scan. Do it **once per trailer, cached** —
   never per statement.

## The exit slots are typed — take the free win

The 25 exit handlers are split by return type: `lblEX_ExitProcI2` / `I4` / `R4` /
`R8` / `UI1`, `ExitProcHresult`, `ZeroRetVal`, `ZeroRetValVar`, plus the untyped
`lblEX_ExitProc`. **Which slot fires tells you the return type before any value is
decoded**, half the return-value answer for nothing. Three members do not end a
frame: `ZeroRetVal` and `ZeroRetValVar` (1662, 1663) release a String or object,
or clear a Variant, just before the real end of a class Function, and a GoSub
`Return` (620) returns within the procedure.

(`lblEX_ExitFor*` are *loop* exits, not procedure exits. Do not treat them as
frame pops.)

## The hook's lifetime rules, learned the hard way

Three rules the VBA hook path obeys:

1. **The thunk saves floating-point state** — `xmm0`-`xmm5` and `MXCSR` — around
   the C++ call, not just the integer registers.
2. **Nothing resets shared state while a thread is inside a hook.** Disarm
   drains an in-hook counter first, and *skips the teardown* if it does not
   drain.
3. **A stack overflow is not an ordinary fault.** Restore the guard page with
   `_resetstkoflw`, open the breaker immediately, and count it separately from
   ordinary hook faults.

Why each holds:

**1.** A handler is entered by `jmp`, not called, so the x64 volatile-register
rule does not apply: the interpreter spilled nothing. The hook clobbers `xmm`
for certain (`_snprintf_s` takes a `double` in `xmm0`; `memcpy` uses `xmm`), and
an unsaved register surfaces as a silently wrong number in the workbook.

**2.** Calc threads that entered a stub just before disarm may still be reading
the procedure table, the totals and the p-code lengths. `WaitForHooksQuiet()`
drains an in-hook counter first, and if it does not drain the teardown is
skipped: stale counters are recoverable, resetting them under a live thread is not.

**3.** Catching `STATUS_STACK_OVERFLOW` like any other fault would leave the
thread's guard page consumed, so the next stack growth on that thread could not
recover. The hook runs at exactly the depths where least stack remains, because
VBA recursion runs until VBE7 gives up; so the guard page is restored, the
breaker opens at once, and `stackOverflows` is counted apart from `hookFaults`.

## The calling cell, from inside the interpreter

A p-code trailer is per-*procedure*, so ten cells calling one UDF are
indistinguishable inside the interpreter. **The cell is the only discriminator,
and it does not exist in VBA's state at all** — it exists in the calculation
engine, so it is asked for rather than derived.

`xlfCaller`, called from inside the dispatch hook, answers. That it answers at
all was measured rather than assumed. It returns an `xltypeSRef`, which carries
no sheet; `xlSheetNm` on that same reference supplies the sheet as a name, which
is what the row wants anyway.

This is a call **into** Excel from a traced thread, so it is not free and it is
not done per statement — only where an activation needs a cell.

**Where the cell cannot be established, the row carries none.** Excel runs
activations of its own with no calling cell, and a confident wrong cell is worse
than an empty one.

## The arguments

**The arguments of a UDF that a cell invoked directly are not readable at the
moment the activation opens** — nothing has pushed anything yet, and there is
nothing to read. Four routes were measured and rejected before the one that
works.

**The callee loads its own parameters, and the load says which is which.** A
load's frame offset is **negative for a local and positive for an argument**, so
taking the positive-offset loads in ascending order *during* the activation
answers the question that reading at the open could not. Capture is therefore
deferred from open to during, and needs no mechanism beyond the walk already
running.

`CaptureArgs` (`vba/vbaargs.cpp`) walks the trailer from `R14`, resolves each
parameter's declared type from the opcode that touched its slot, and renders the
value. An opcode that touches a parameter slot but is not in the type table
renders the declared type as `?opNNN` rather than guessing, and is counted.

## The return value: the exit opcode says the type, the store says the rest

**There is no single frame slot that always holds a return.** `[R14−8]` holds
one for some Functions, the first local for others and a compiler temporary for
others again — measured across thirteen procedures. Reading that slot
unconditionally is what produced a Sub reporting `ret='1' rettype='Long'`. The
type is therefore established before the slot is read, from two sources, and the
row is left **empty** when neither answers.

**1. The exit opcode.** Most typed Functions leave through a typed exit handler,
and `ExitReturnKind()` maps it — 623–635 and 952 for the scalars. An exit opcode
that maps to `RetKind::None` means the procedure has no result: a Sub, or a
Property Let. **504 is one of these** (the class/form Sub exit), and mapping it
was what stopped a Sub's local being reported as a return.

**2. The store opcode, when the exit is type-agnostic.** Class, form and Property
procedures leave through **1664** (Function) and **504** (Sub), which carry no
type. For those the trailer is scanned backwards for the store into the result
slot, and the store opcode names the type — the store opcodes are the load
opcodes plus 32. Two known stores of *different* types are a **refusal, not a
pick**: one is a false positive inside another instruction's operands and nothing
available can say which.

**A Variant is identified by its slot, not by its opcode.** A Variant result
lives at `[R14−0x18]` and nothing else does. Measured, VBA stores a Variant
through *different* opcodes depending on what it holds — 694 for a number or an
array, 707 for a string — so enumerating opcodes would mean a new one for every
held type, each discovered only when somebody happened to return it. The offset
is the invariant, so the offset is what is trusted; `DescribeReturn` then reads
the `VARIANT` tag for the held type.

**Two exit opcodes fire mid-activation and are not endings.** 1662 and 1663
release a temporary for String/Object and Variant results respectively, *before*
the result is stored. Treating them as endings closed the frame early, emitted an
empty `ret`, and then re-opened the procedure — one call, two entry and two exit
rows. They are on `ProcedureEnd`'s deny-list. The regression that catches this
has VBA count its own calls and compares (`class-returns-every-type`); a trace
that disagrees with the interpreter about *how many times a function ran* is a
worse defect than a missing return value, and only that comparison sees it.

Measured coverage: every scalar, `String`, `Boolean`, `Object`, `Variant`
holding a scalar, an array or a string, out of standard modules, class modules,
forms and Property Gets — with Subs and Property Lets correctly carrying
nothing.

## A parentless row is legitimate

Excel runs a pass over user functions about a second after a recalculation
finishes, once per VBA cell, with no calling cell. Those rows carry neither cell
nor sheet. The model must be able to represent that rather than invent a caller.

## The outcome column: where an error went, and where it escaped

Every VBA exit row carries an `outcome` from a closed set — `returned`, `threw`,
`unwound`, `handled`, `abandoned`, `unhandled`. The first five describe an error's
path through the shadow stack; `unhandled` marks where an error left VBA into a
worksheet cell. This is the most subtle mechanism in the tracer, so the reasoning
is set out in full.

**The chain, for an error that stays inside VBA.** The raise opcode (497) fires for
a real `Err.Raise` and for ordinary object-model VBA alike, so nothing *at* the
raise says whether an error is real. The frame's fate does, and is known only when
it closes: `DecideOutcomeAtClose` (`vbatrace.cpp`) reads how the frame left. A frame
that raised and ran no epilogue **threw**; each outer frame the error passes through
without resuming is **unwound**; the frame that runs again after the raise
**handled** it. So a chain reads `threw → unwound → … → handled` outwards from the
raiser. A raise that ran its own epilogue is a benign object-model raise or a
same-frame handler and reads `returned`. This part is exercised by
`tests/sweep/vba/error-shows-thrower-and-catcher` and `error-from-class-and-form`.

**The problem this section solves.** An unhandled error in a worksheet function does
not propagate anywhere a user can see as an error — Excel turns it into `#VALUE!` in
the cell. The demo `01_CalcChain.xlsm`, run without its helper add-ins, is the case:
`OptionBook` in a cell fails, the cell shows `#VALUE!`, and a macro that recalculated
the sheet carries on. The trace must say three things: the function `threw`-then-left
(a distinct outcome, `unhandled`), the recalculating macro `returned` (it never saw
the error), and — separately — a function that *returns* `CVErr(xlErrValue)` reads
`returned` with the error in `ret`, because it did return. That last distinction,
returning an error value versus failing with one, is why `unhandled` is a separate
word rather than reusing `threw`.

**The hard part: who started this frame.** Marking `unhandled` needs to know, when a
frame opens, whether Excel entered VBA to compute a worksheet function (an error
escapes to the cell) or VBA called the frame itself (an error propagates to the VBA
caller). Both cases look almost identical from the interpreter: the parent's
statement is mid-execution, and both may reach the interpreter through COM. The stack
alone conflates them. Four approaches were tried; the reasoning is worth keeping
because each looked plausible.

- **Mark `unhandled` only when the error escapes every VBA frame (reaches the bottom
  of the shadow stack).** Simplest, and correct for a standalone failing cell. But it
  loses the demo: a cell function failing under a recalculating macro never reaches
  the bottom, because the macro is still beneath it, so the macro would wrongly read
  `handled`.
- **A stack-step size threshold.** A direct VBA-to-VBA call lowers the interpreter's
  `rsp` by exactly `0x160`; Excel entries step further. But the wider steps do not
  separate by size — a class method call steps `0x1660`, a cell recalc under a macro
  `0x76b0`, and a class method with more arguments would cross any threshold. It broke
  `error-from-class-and-form`: a class method's error stopped at the class instead of
  reaching its VBA caller.
- **Scan the stack gap for an `EXCEL.EXE` return address.** Read the stack between a
  frame and its parent; an Excel return address means Excel's code sits between them.
  It is fooled by *stale* return addresses: passing a `Range` to a VBA method leaves a
  call-preceded `EXCEL.EXE` address in the gap from evaluating that argument, so a
  plain VBA method call read `unhandled`. Pinned as a rejection by
  `an-excel-object-argument-does-not-fool-the-boundary`.
- **Walk the live return chain with `RtlVirtualUnwind`.** Immune to stale addresses,
  because it follows only real frames. But our own per-slot stub sits in a
  `VirtualAlloc` page with no unwind data, so a walk starting inside the hook cannot
  step past it to reach the interpreter; and it would mark an `Application.Run` macro,
  which is Excel's own code, as an escape when that call actually propagates. It also
  reintroduces the process-terminating stack-walk hazard (Part 5).

**What is used: the calling cell, compared to the parent.** Excel enters VBA to
compute a worksheet function, and `xlfCaller` (the C API, `Excel12(xlfCaller)`, not
COM `Application.Caller`) names the cell it is computing. A VBA call within that calc
keeps the same cell, so a frame whose calling cell **differs from the frame beneath
it** — or has a cell where the frame beneath has none — is a fresh worksheet-function
entry, the one activation whose unhandled error becomes `#VALUE!`. `IsExcelTheCaller`
(`vbatrace.cpp`) is that comparison: it hashes the caller cell onto the frame
(`Frame::callerHash`) and compares to the parent's. `EndErrorAtExcelEntry` marks such
a frame `unhandled` when an error leaves it and ends the chain there, so the frame
beneath reads `returned`. It is the simplest of the four — one lookup, already taken
for the `caller` column, and one comparison — needs no stack walking, and gets every
case right: cell functions (`unhandled`), class and property calls and object
arguments (propagate), `Application.Run` (propagates, which is what Excel does), and
composed formulas like `=Test(test2(),test2())` where each call is its own cell entry.
Pinned by `error-into-a-cell-stops-at-the-cell` and
`a-function-can-return-or-raise-an-error`.

**Because the comparison needs the calling cell on every frame, resolving it is no
longer optional** — the `CALLER` setting was removed. The lookup costs about 3.5 µs
per traced call and is now unconditional.

**The one gap.** An activation Excel starts with *no* calling cell — a user-triggered
sheet event, an `Application.OnTime` macro — is not seen as a cell entry, so its
unhandled error reads `threw` rather than `unhandled`. It is under-labelled, not
wrong: the raiser is still located, the escape is counted, and Excel shows its own
modal dialog. A *VBA-triggered* event (a macro writes a cell, firing
`Worksheet_Change` beneath it) is not a gap — the error genuinely propagates to the
macro, and the trace shows that. Both are pinned by
`a-user-triggered-event-error-reads-threw`.

**Two supporting pieces make the chain hold up.** A frame killed by an unhandled
error runs no epilogue, so the stack-pointer backstop never closes it; the next cell
Excel calculates would nest under the dead frame. Before any new activation opens,
`CloseFramesThatLostTheirStack` closes frames whose trailer is no longer at their
recorded dispatch pointer (`FrameStillLive`). And disarm must not read a frame it
closes while still running as a throw: `Application.Run` fires the raise opcode with
no error, so a macro still executing at disarm keeps its outcome — the same
`FrameStillLive` check tells a running frame from a leftover of an earlier unwind.

---

# Part 5 — Safety

Applies from the first hook installed and is not negotiable.

## Never modify the workbook

- No cells changed. No file written to the user's workbook.
- **No VBA source read at all.** The VBA work observes the interpreter; it does
  not read source code.
- Calibration never uses the user's sheet.

## Fault isolation

- **SEH-wrap every hook body** so a bad interaction cannot take down the host.
- **VBA circuit breaker:** after 8 hook faults in an arming session, or one stack
  overflow, every hook returns at once and the tracer stands down. The table stays
  patched until disarm, which is reversible — nothing is unpatched from an
  arbitrary thread mid-dispatch.
- **XLL recorders have no breaker.** A fault while decoding loses that row, is
  counted and reported at disarm, and the traced call itself still runs.

## Fail closed, then degrade loudly

**Arming is all-or-nothing per mechanism.** The VBA tracer does not arm unless
its structural checks passed and `verified` is set. The XLL tracer arms per
function and declines the ones it cannot parse, resolve or hook — a decline is a
counted outcome, never a silent one, and each has its own counter:
`typeTextUnparsed`, `procNotFound`, `moduleNotLoaded`, `stubSpaceExhausted`,
`ownModule`. The disarm line reports them by name, because "we armed 46 of 53"
is only useful beside *why* the other seven were refused.

The two sides fail differently because they derive differently, and the rule that
covers both is the same: **an address that was not established is not used.**

A session always surfaces what it armed and what it declined, in the log and to
the user. **A degraded session is a normal outcome, not an error** — but it must
never be silent, because a reader has to know whether "no rows" means fast code
or a failed lock.

## CET shadow stacks, and the build flag that must stay off

The XLL exit edge takes a genuine `call` into MinHook's trampoline and regains
control on the return, so the shadow stack records our push and matches it on
`ret`. **CET is satisfied by construction**, not by an exemption.

The rule that remains load-bearing is a negative one:

> **`XRayXL64.xll` must NOT be built `/CETCOMPAT`.** Shadow stacks are enabled per
> process from the main executable, and `EXCEL.EXE` carries no `CET_COMPAT` bit;
> the default compatibility mode forgives mismatches whose return address lies in
> a module not built `/CETCOMPAT`. That flag is the obvious hardening a security
> review will ask for, and adding it would remove the exemption that keeps any
> planted address legal. **Marking our own binary as more secure is what would
> break the tracer.**

**No current Excel is affected.** The bit is spreading through Office — 14 of
719 x64 binaries in the Office tree carry it, among them
`PowerPivotExcelClientAddIn.dll`, `ACECORE.DLL` and `xmsrv_xl.dll` — but none of
them is `EXCEL.EXE`, and a DLL's bit does not turn shadow stacks on for the
process. Breaking the tracer would take all three of: `EXCEL.EXE` gaining the
bit, the process running in strict rather than compatibility mode, and an exit
edge that rewrites a return address, which neither tracer uses. The trend is
why that last one was engineered out, and why the rule above stays.

## AV / EDR

This code hooks Excel internals; expect AV/EDR to notice. The XLL vehicle rather
than an injector is a deliberate part of the mitigation. Signing, allow-listing
and security-team engagement are deployment concerns, not current-scope concerns,
but they have long lead times and are already overdue.

## x64 first

Target 64-bit Excel only. **The decoders are x64 throughout** — arguments read
out of the x64 register file, `XLOPER12` at 64-bit widths, the interpreter's
frame reached through `R14`. 32-bit is not a build flag; it would be a second set
of decoders, and it is out of scope.

---

## Known limits of what is built

Properties of the code as it stands, not a forward plan.

| | Limit | |
|---|---|---|
| 1 | **Cross-version durability is argued, not tested.** Every address is derived, which is argued to survive an Excel update by construction — and the dispatch-slot indices were checked across 41 VBE7 builds from 2012 to 2026, but the tracer as a whole runs only against the Excel its tests run on. | The largest long-term risk. |
| 2 | **Silent mis-derivation produces complete-looking rows.** A decoder reading the wrong location does not fault; it fills the row. | Mitigated by the suites' planted answers and by decline counters. |
| 3 | **Coverage is bounded by enumeration.** A function hook sees exactly what was enumerated and armed; a function that could not be resolved or hooked, or that registered while the registration watch was off, is invisible rather than merely untraced. | This is the price of zero derived addresses. Every such case is a counted decline, and it must never be silent. |
| 4 | **The dispatch table is shared.** Patching it affects every VBA consumer in the process, and a second add-in doing the same thing would collide. | Reversible — disarm is a pointer write back, checked against the originals recorded at arm. |
| 5 | **Arming cost scales with the user's add-in, not with what is traced.** | Measured below, in Part 3. |
| 6 | **The observer effect has never been measured.** Instrumented versus un-instrumented total recalculation time. | Outstanding since the first spike. It should be a number before anyone is asked to run this on a machine they depend on. |
| 7 | **Excel-DNA identity is an export name.** `f0`, `f1` … tell the reader nothing, and tell it confidently. | Resolved at arm time through the `xlfRegister` hidden-name / `xlfGetDef` chain; measured resolving 46 of 46. |
| 8 | **The VBE Reset button is untested.** It is a UI action inside the Visual Basic editor with no object-model equivalent, so it cannot be driven from a script the way `End` can. | Same class of teardown as `End`, which passes — an argument by analogy rather than a measurement. |

---

## Acknowledgements & prior art

The VBA p-code tracing technique — locate the `vbe7.dll` opcode dispatch table by
its structural fingerprint, patch opcode slots, recover procedure identity from
interpreter state, time with `QueryPerformanceCounter` — is inspired by and
validated by [Dr. D. Azzopardi's `desva/VBA` project](https://github.com/desva/VBA).
This design adopts the *approach* and builds an independent implementation; it
reuses none of that code.

**And the standing lesson from that fact:** the p-code argument layout was
documented in published work the entire time it was being searched for in memory.
**Search the literature before deriving from scratch.**

---

