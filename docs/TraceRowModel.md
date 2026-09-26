# The trace row model — the contract for XRayXL's output

What the trace file contains, guaranteed, and what a reader may rely on. If
you only want to *read* a trace, the README's tour is enough; this is for
anyone writing a reader, a test, or a change to the writer. The settings that
decide what gets recorded are in [TraceOptions.md](./TraceOptions.md). The
writer is `src/emit/csv.cpp` and `src/emit/rowcsv.cpp`, plus the two row builders
(`src/xll/xlltrace.cpp` and `src/vba/vbatrace.cpp`); the
enforcing reader is `Read-TraceFile` in `tests/sweep/_xray_common.ps1`.
If the code and this document disagree, the disagreement is **loud** — the
reader refuses the file — and the fix updates both in the same change.

## Three output channels, one invariant

| Channel | File | Nature |
|---|---|---|
| **Trace** | `%TEMP%\XRayXL\TraceFiles\XRayXL_Trace_<id>_<pid>.csv`, or `.jsonl` with `FORMAT=JSONL` | The data product: one row per event. `<id>` is a monotonic OS tick (`GetSystemTimePreciseAsFileTime`) that always rises — across arms, processes and reboots — so a re-arm never overwrites; created lazily on the first record, which is the session's `arm` row, so every arm leaves one: an arm that armed nothing leaves its `arm` and `disarm` rows. Buffered; drops shown by `input`-column holes |
| **Log** | `%TEMP%\XRayXL\Logs\XRayXL_<pid>.log` | Control: arm outcomes, the derivation line, commands, and the disarm report — the totals line and the procedures line. One line per message. Appended synchronously, **never dropped**. Levelled (DEBUG/INFO/WARNING/ERROR), default INFO |
| **Crash log** | alongside the logs | Notes that must survive the process dying; opened `FILE_APPEND_DATA` per note, installed with the top-level exception filter |

Totals are counted **in memory at capture** (interlocked, no I/O) and written
to the *log* at disarm — they never pass through the trace file's write
path. That separation is what makes loss checkable: **totals say what
happened, rows say what was recorded, and the difference is reported out of
band — a loss count on the control channels — and located in the file through
the `input` column's holes**. If
totals were derived from the file, drops would silently shrink them and the
check would be vacuous.

## Lifetime

- One file **per arming session**, named with the monotonic `<id>` and the
  pid. A re-arm writes a **new** file with a new `<id>` and never overwrites
  the last.
- Created lazily on the **first record** (with `CREATE_ALWAYS`), the session's `arm`
  row: a file holds exactly **one arming session**, and `seq` restarts at 1 in each
  new file. Reading after a re-arm reads the new session only — this is the
  suites' isolation guarantee.
- Flushed and closed at **disarm** (`FlushFileBuffers`). Reading after your
  own disarm sees the complete session; a mid-arm read sees a prefix.
- Disarm additionally **drains** any buffer before returning. The
  read-after-disarm completeness guarantee is the part of this contract that
  buffering must not change.

## The header

Byte-for-byte, the first line is:

```
seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,caller,callerref,argcount,args,ret,rettype,outcome,ticks,tracerticks,trust
```

Encoding: UTF-8, CRLF line endings, minimal RFC 4180 quoting — a field is
quoted only when it contains a comma, a quote or a newline, and embedded quotes
are doubled. Module names, workbook names and addresses are written as UTF-8.
Inside `args` and `ret` every character above 126 is escaped as `\uNNNN` and a
line break as `\r` or `\n` (see *String escaping*), so a value is always ASCII;
a CR or LF anywhere else becomes a space. A row is at most 4 × 4 MB + 64 KB, the
two value columns at their limit; a row larger than that is dropped whole, and
its `input` value is a hole.

**One optional column.** With `XRayXL_SetTraceParam "VBA", "BREAKPOINTS", TRUE`
set when tracing arms (and VBA traced), the header ends `...,ticks,tracerticks,trust,breaks`
instead, and every row of that file has the extra field. Nothing else ever
changes the header: a reader accepts exactly these two, and refuses any other.
In JSON Lines the key simply appears on the rows that carry a count.

## Kinds and sources

**What** happened is `kind`; **who** it happened in is `source`. They are two
columns because they are two facts. One column carrying both (`vba-entry`)
forces a filter on either to spell out the other, and makes "every entry row" a
wildcard match.

| `kind` | `source` | Meaning |
|---|---|---|
| `entry` / `exit` | `XLL` | A hooked add-in function activated / returned |
| `entry` / `exit` | `VBA` | An interpreter frame opened / closed |
| `event` | `Excel` | Excel raised one of its Application events, named in `function` |
| `event` | `XRayXL` | Tracing started (`arm`) or stopped (`disarm`) |

**Every trace starts with the `arm` row and, when tracing stops cleanly, ends with the `disarm`
row.** `arm` carries what a reader needs to read the rest: `qpcFrequency`, the `utc` time its `qpc`
was taken at, the `pid`, and every setting in force. `disarm` carries the session's totals
(`rowsDropped`, `eventsRecorded` and the rest), each `0` for a source the session did not run.
Both are `name=value` in `args`, with no type, and only text quoted:
`qpcFrequency=10000000 utc="2026-09-25T15:13:00.0910753Z" pid=52204 XLL_DEPTH="ALL" XLL_ARGS=TRUE`
and `rowsDropped=0 vbaStoodDown=FALSE`. A trace with no `disarm` row stopped because
Excel did, by a crash or a hang.

**Excel's events** are recorded in the order Excel raises them, which is not always the order one
might expect: an edit reads `SheetCalculate`, `AfterCalculate`, then `SheetChange`, because Excel
recalculates before it raises the change. Which events are recorded is the `EVENTS` setting, one
event at a time, the Calc set by default:

```
XRayXL_SetTraceParam "EVENTS", "SheetSelectionChange", TRUE
XRayXL_GetTraceParam "EVENTS", "SheetSelectionChange"     ' TRUE or FALSE
XRayXL_GetTraceParam "EVENTS"                             ' the preset the set matches, else Custom
```

An event row has `span`, `parent` and `depth` `0`, since an event is not a call. `callerref` is
the most specific place it names (the target range, else the sheet, else the workbook), quoted as
a cell's is (`'[has space.xlsx]Sheet1'`), and `args`
holds its parameters by name: `Sh:Object=Worksheet@0x...([Book1]Sheet1) Target:Range=Range@0x...([Book1]Sheet1!A1)=5`.
An event a newer Excel adds, which this add-in does not know, is recorded under the name its type
library gives it, when every known event is chosen.

There is **no drop-marker row**: the CSV holds
only real events. A `BUFFERWHENFULL=DROP` run that lost rows reports the count OUT OF
BAND -- returned by `XRayXL_Disarm`, on the status summary, and on the disarm log
line. The file still shows loss, but through the **`input` column's holes**,
not a phantom row: a gap in the producer's `input` sequence is exactly where a row
was dropped. So the count comes from the operator's channels and the *location*
from the file itself.

Both sets are **closed**: a new kind or source arrives only together with an
update to this document and to the reader, so an old reader meeting a new
value refuses loudly rather than mis-filtering silently.

Which rows *appear* is mode-dependent:
`XRayXL_SetTraceParam "VBA","DEPTH","TOP"` emits only depth-1 VBA frames
(`"ALL"` emits every frame; the totals count everything either way), and `XRayXL_SetTraceParam "XLL","DEPTH","OFF"`
means no XLL rows because nothing was hooked. The row *shapes* above are
mode-independent.

## Columns

| Column | Type | Contents per kind — and what EMPTY means |
|---|---|---|
| `seq` | int64 ≥ 1 | Stamped by the **writer** (the drain, or the sync producer under the lock): seq order **is** file order, across both sources. Strictly increasing and **dense** — every written row gets the next number, so a lossy file still reads `1..N` (which is why drops are not visible from `seq`) |
| `input` | int64 ≥ 1 | Stamped by the **producer** at emit: unique, but **not** contiguous and **not** in file order. A dropped row consumes an `input` value that never lands, so its **holes show where, and how much, data was dropped**. Sort by `qpc` (then `input`) for event order; the holes are the in-file record of loss that replaced the `gap` row |
| `kind` | enum | `entry`, `exit`, `event` -- see above |
| `source` | enum | `XLL` or `VBA` for a call; `Excel` or `XRayXL` for an event |
| `span` | uint64 | Pairs an entry with its exit. **One sequence for both sources**, increasing in the order calls started, and unique for the life of the Excel process. It does not restart when tracing is armed again, so a later trace file's spans need not begin at 1 — a call still running across a re-arm can never reuse a number |
| `parent` | uint64 | **Every row, never empty.** The `span` of the activation of the same source that called this one; `0` at the top of a chain, which is a value and not an absence |
| `depth` | int | **Every row, never empty.** How far down the call chain this activation was, counted per source and per thread; `1` when no other frame of the same source was open |
| `thread` | uint32 | OS id of the traced thread |
| `qpc` | int64 | `QueryPerformanceCounter` at the **event** — never wall clock. A row is *stamped* when it happened and *numbered* when written; `seq` and `qpc` order different things |
| `module` | text | XLL: the add-in file (`PricingLib.xll`). VBA: the qualified module (`[Book.xlsm]Module1`, and `[unsaved:006cd206e1]Module1` for a workbook that was never saved, which has no file name for VBA to report) — or **empty when unresolvable**, never a guess like `VBE7.DLL` |
| `function` | text | XLL: the registered name. VBA: the resolved procedure name, falling back to the trailer address in hex — an address obviously *looks like* an address |
| `proc` | text | XLL: the exported procedure name. VBA: **always** the trailer in hex — the identity the tracer actually used, which tells two same-named procedures apart |
| `typetext` | text | **Entry rows only.** The parameter types, comma-separated, with no brackets: the XLL registration's argument codes (`B,B`, `O%`, `D%,K%,E`) or the VBA declared types (`Long,String`). **Empty** when there are no parameters, where `argcount` reads `0`; and when VBA could not recover the types, where `argcount` holds the slot count or is empty. Empty on exits |
| `caller` | text | Entry rows only, and **never empty**: what the caller *was*, as a kind from a closed set — `cell`, `name`, `toolbar`, `menu`, `editor`, `registerid`, `none`, `unavailable`, `array`, `unknown`. Two of those are different "no" answers and are kept apart deliberately: `none` is Excel saying there is no caller on a sheet, `unavailable` is Excel declining to answer. The calling cell is always resolved, so a row is never left unasked. `name` rather than `object` because a graphic object and an `Auto_*` macro both come back as a string and Excel gives no way to tell them apart — the kind says what is certain |
| `callerref` | text | The description, whose meaning `caller` decides. `cell`: one **external address**, `[Book1]Sheet1!B2` or `'[my book.xlsx]Sheet 1'!B2` when Excel would quote it, or a whole range `…!B2:D4` for a CSE array formula — the same text `Range.Address(,,,True)` returns, so it can be pasted back. `name`: the shape's name, or an `Auto_*` macro's calling sheet. `toolbar`: the button's position, then the bar: `2/5`, or `2/"MyBar"` when a custom bar answers with its name; the Quick Access Toolbar is bar `0`. `menu`: the command's position, the menu's, the bar's ID and the submenu's (`0` for none), so `27/27/14/0` is the 27th command on the cell's right-click menu. Excel sends both in the reverse of the order its documentation gives. `editor`: empty; the VBA editor started it (Run, F8 or the Immediate window). `registerid`, `unavailable`, `array`, `unknown`: the one value. `none`: which no — the short name of the Excel error it answered with (`ref`, `value`, `name`, …, or `err<N>` for one without a name), `nil`, `emptyref`, `sheetless-B2`, `nametoolong` |
| `argcount` | int | **Entry rows only.** The number of parameters. XLL: one per type code, so an `O` array counts once though it takes three slots. VBA: the parameter count when the types were recovered, the slot count when only that was, empty when capture did not check out |
| `args` | text | **Entry rows: the values going IN. Exit rows: the ByRef ones that CHANGED** — a VBA exit row carries `args` only when a re-read at the exit differs from the entry, so a row that has them is saying "these moved"; absent means unchanged, and `byrefEligible`/`byrefChanged`/`byrefSame`/`byrefDeclined` in the disarm line separate that from "never looked". `argcount` and `typetext` stay **entry-only**: they describe the signature, which has not changed, and the entry row this pairs with by `span` already carries them. **ByVal is never reported at an exit** — the callee's copy may differ but the caller never sees it, so an "after" value would assert an effect that does not exist. One field, `a<N>:<type>=<value>` per argument on both sources — `a<N>:<type>@0x<address>=<value>` where a VBA argument's storage has an address to match on (see *Reading `args`*) — joined by single spaces — one column whatever the arity, which keeps the file rectangular. On XLL rows `N` is the argument's position. **On VBA rows `N` is the argument's first frame slot**: a `ByVal Variant` fills three slots, so the argument after one reads `a4`, and `argcount` still counts arguments, not slots. **A VBA `Variant` argument is decoded by the same code as `ret`**, so it reads its held value the same way, by the rules in *Values*: a Double bare (`1.5`), any other number named (`Integer(42)`, `Decimal(12345.678901234567890123456)`), `"text"`, `TRUE`, an array (`Variant[1..2,1..2]{{1,"x"},{2,TRUE}}`), an object (see *Objects*), `Nothing`, `Empty`, `Null`, and an Excel error **spelt as Excel spells it** — `#N/A`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#NULL!`, `#GETTING_DATA` — falling back to `Error(0x…)` for an SCODE outside that set. An omitted argument reads `Missing` on both sources. A slot that decodes to nothing truthful is the raw qword (`0x…`), never a coerced value. **On VBA rows the type is the declared type where VBA's metadata named one, and a `?` marker where it did not** — see *Declared or inferred* below. On XLL rows it is the registration's code, and a value with nothing to decode reads as a bare word: `Missing`, `Empty` or `AsyncHandle`; a reference argument (`R` or `U`) reads `SRef(R2C2:R3C3)` for one area on the calling sheet and `Ref(R2C2:R3C3,R5C5:R5C5)` otherwise, the sheet not named; anything else `?xltype<N>` |
| `ret`, `rettype` | text | **Exit** rows only. XLL: the decoded return value, and in `rettype` the registered return code followed by any registration flags — `Q`, `Q$` (thread-safe), `Q!` (volatile), `Q#` (macro-sheet equivalent), `Q&` (cluster-safe). `ret` is empty for a function registered with no return value (`>`, or a modify-in-place digit) and for an async call, whose answer arrives later. VBA: every **Function** exit at any depth, `rettype` naming the kind the exit opcode declared — `Double` (also Date), `Single`, `Byte`, `Integer` (also Boolean, as −1/0), `Long`, `LongLong`, `Currency`, `String` (quoted), `Object`, `Variant`, `Elem()` for a typed array, or `Udt` for a user-defined `Type`, whose `ret` is its address, `udt@0x…`, as a record argument reads. An array, a Variant and an object read as they do in `args`, by the rules in *Values* — `Long[1..3]{3,6,9}`, `Double[0..1,0..1]{{1234.5,2},{3,4}}`, `Long(777)` for a Variant holding a Long — so one object can be followed from argument to result. **Empty** for a Sub (no result exists) and anything whose descriptor fails validation |
| `outcome` | text | **Exit rows only, and never empty on one.** How the activation ended, from a closed set: `returned`, `threw`, `unwound`, `handled`, `abandoned`, `unhandled`. An XLL exit row always reads `returned` — see below. Empty on every other row |
| `ticks` | uint64 | **Exit rows only.** The duration in QPC ticks, spelled the same by both sources. Empty on the one exit that has no duration — an async XLL call, where the span measured the dispatch and the work has not finished |
| `tracerticks` | uint64 | **Exit rows only, beside `ticks`.** How many of `ticks` the tracer's own hooks spent inside this activation: reading arguments, asking Excel for the calling cell, writing rows, and waiting on a full buffer under `PAUSE`, for this call and every call it made, on both sources. `ticks` minus `tracerticks` is the call's own time. The per-statement VBA check is too small to time and is not included. Empty whenever `ticks` is |
| `trust` | text | **Exit rows only, and never empty on one.** What ended the measurement, and so whether `ticks` is a reading or a ceiling: `exit`, `end`, `backstop`, `flush`, `async`. `exit` means the same thing on both sources — the return path fired. **`async` is NOT SUPPORTED** — see below |
| `breaks` | uint32 | **Optional — only in a file traced with `BREAKPOINTS` on.** On every VBA exit row, how many times that activation paused in the VBA editor -- at a breakpoint, or at a `Stop` statement in the code; `0` when it did not. A non-zero count means `ticks` includes time spent in the debugger, not running. Its callers' `ticks` include that time too, though their own count is `0`. Empty on every other row |

**`outcome` — how an activation ended.** A value from a closed set that readers
group and count by, which is why it is a column: a substring match on free text
can only approximate one. One of `returned`, `threw`,
`unwound`, `handled`, `abandoned`, `unhandled` on every `exit` row. A chain reads outwards from the
throw: the frame that raised says `threw`, each frame the error passed through
without handling says `unwound`, and the frame that resumed says `handled`. A frame
that caught an error and then raised one that left it says `threw`, and the frame
that catches that one says `handled` in turn.
`returned` is written explicitly, never omitted, so a clean call cannot be
confused with an older writer's silence.

**The scope is errors that CROSS a frame boundary.** An error raised and
swallowed inside one frame -- a same-frame `On Error Resume Next` -- reads
`returned`, because nothing propagated and the activation did return. `threw`,
`unwound` and `handled` describe a chain, and a chain needs at least two frames;
`handled` in particular means *this frame caught something thrown below it*,
which is the thing a reader is hunting for. Reading it any other way would make
`handled` common and uninformative.

**`unhandled` is an error that left VBA into a worksheet cell.** The function Excel
entered to compute a cell reads `unhandled`, and the cell shows `#VALUE!`. Frames it
called keep `threw` and `unwound`, so the raise is still located, and the chain ends
there even when a macro recalculated the cell, which then reads `returned`. The
frame is identified by its calling cell (`xlfCaller`, always resolved) differing
from the frame beneath. An entry with no calling cell, such as a user-triggered sheet event, an
`Application.OnTime` macro or an `Application.Run` macro, reads `threw` instead. A function that *returns* an error value, such as `CVErr(xlErrValue)`,
raised nothing and reads `returned`, with the error in `ret`. A frame still
running when disarm closes it, such as a macro that disarmed, never reads `threw` or
`unwound`.

**`abandoned` is the fifth, and it is not an error.** `End` tears the whole VBA
session down without running an epilogue anywhere, so the frames it kills
neither returned nor threw. They are marked `abandoned` -- never `returned`, which would
assert a return that did not happen.

**Two things give it.** The `End` statement fires the `End` opcode, which closes
every open frame at once, `abandoned` with `trust=end`. The **End** button on VBA's
error dialog fires no opcode, so its frames are closed later, when the next VBA call
finds them all gone (`trust=backstop`) or at disarm (`trust=flush`). There the
innermost frame raised the error and reads `threw`; the frames beneath read
`abandoned`, and frames from a cell entry upwards keep the cell's `unhandled`. Every
case, with examples, is in [ErrorsInTheTrace.md](./ErrorsInTheTrace.md). Two things that sound like they belong here do not. Excel
rescheduling a formula whose precedent was not yet calculated does **not**
abandon the call -- the activation runs to completion and the UDF is simply
called *again*, so the trace shows two whole `returned` pairs, not a truncated
one. And a user break during a long calculation is *cooperative*: an XLL is
merely asked, through `xlAbort`, and returns of its own accord; VBA gets error
18, a trappable error like any other, which reads `threw` or `handled`. Excel
has no way to take a running activation away, which is why there is no third
route.

**An XLL exit row always reads `returned`.** The row is written only on the
call's return path, so the word is always true; the other four describe an error
chain, which only VBA's shadow stack can follow. An XLL call that ends in an
exception writes no exit row at all. `trust` `async` is a separate statement,
about the number beside it: an async function returned perfectly normally, it
just returned before its answer existed.

It holds for class modules, property accessors, constructors and form modules
as well as ordinary procedures, and each of those is a separate test -- an
error raised inside `Class_Initialize` is attributed to the constructor, not to
whoever said `New`. **Two limits.** A raise inside `Class_Terminate` is not
propagated by VBA at all (it puts up a message box instead), so nothing on that
chain is marked `handled`. And the error's NUMBER and DESCRIPTION are not
recovered: `outcome` says WHERE an error was thrown and who caught it, never
WHICH error it was.

### Declared or inferred: how far to trust a VBA argument

Every VBA argument reads `a<N>:<type>=<value>`, and the `<type>` says how the value was
found. **A named type means the value is certain. A `?` means the tracer had no type and
recognised the value from the bytes themselves** — a strict recognition, but a recognition.

```
a1:Long=42                         declared Long; the slot was read as a Long          certain
a1:Variant=Long(42)                declared Variant; the Variant's own tag says Long   certain
a1:Ref&=Long[1..5]{4097,4098,…}    declared only "a reference"; the array recognised   recognised
a1:?none=Double[0..400]{1,0.99,…}  no type; the bytes proved to be an array of Double  recognised
a1:?unseen="EUR"                   no type; the bytes proved to be a string            recognised
a1:?none=Missing                   no type; an omitted Optional, two exact constants   recognised
a1:?none=0x4004000000000000        no type, and nothing recognised: the raw 8 bytes    raw
```

**With a type**, the slot is read *as* that type and nothing else. A `Variant` is
self-describing in a precise sense: VBA stores the held value's type in the Variant
itself, so once the declaration says the slot is a Variant, its tag says the rest. A
declared value never falls through to recognition — a `Double` whose bits happen to look
like a string cannot render as text.

**Without one**, the tracer has eight bytes and no declaration. It tries a few shapes in a
fixed order, each directly or one pointer away, and accepts one only if **every** check
passes:

| recognised as | it must |
|---|---|
| `Missing` | carry exactly the tag and code VBA uses for an omitted `Optional` — two exact constants |
| a string | be a length-prefixed string in user memory, 2-aligned, of a plausible length, ending in a terminator exactly where the length says, with no control characters other than tab |
| an array | be a valid array descriptor: 1 to 8 dimensions, only documented flags, sane sizes and bounds, a data pointer in user memory — **and name its element type**, with the element size matching that type (8 bytes for Double). Something merely *shaped* like an array is refused |

Every array parameter the suites pass — Long, Byte, Integer, Object, Boolean, Single, Double,
Date, Currency, LongLong, String, Variant — names its element type, so the rule costs no real
array. An `Enum` array names `VT_USERDEFINED` and reads as `Long`, which is what its elements
are. The one array that names nothing is an array of a user-defined `Type`: it reads as the
raw qword, and a `Type`'s fields are never walked. The weakest acceptance is a string: a length prefix, alignment, a bound and a
terminator in exactly the right place — strong, but metadata rather than a tag. `Ref&` is
the bytecode saying "eight bytes, by reference" and nothing more, which is why an array
passed `ByRef` is recognised rather than declared.

If nothing passes, the value is the **raw 8 bytes**, `0x…`, and never an invented number.
`0x4004000000000000` above is a Double's bits for 2.5; the trace will not say so, because
nothing told it the slot was a Double.

**How far to trust a recognised value:** almost completely, but it is inference. Eight
unrelated bytes would have to point at something passing every check above. The `?` is
there so a reader knows which values are declared and which are recognised — a wrong value
that still parses satisfies every check, and only the declaration rules that out.

**Why a type is missing.** VBA's bytecode names a parameter's type only at an instruction
that uses it in a typed way. A parameter the body never reads, or reads only by passing it
on, never meets one. The type exists in the source — `ByVal n As Long` — but not in the
instructions the tracer reads.

**Except when it is passed on into a Variant.** Handing a parameter to a `Variant` parameter,
or to a built-in such as `Year()`, makes VBA wrap it in a by-reference Variant, and the
instruction that does so carries the Variant's `VARTYPE`: the callee reads the parameter by it,
so it is the declared type. The tracer reads that label from the instruction straight after the
push, and the parameter is named like any other, `Date` as `Double` and `Boolean` as `Integer`,
as their loads spell them. A `ReDim` of an array parameter names it the same way, as `Ref&`,
from the element type the `ReDim` allocates. A typed instruction anywhere in the body still decides. What the
label cannot reach is a parameter passed on to a typed `ByRef` parameter of another procedure:
there is no Variant, and it stays `?none`.

The word after `?` says which case it was:

| `a1:` | meaning |
|---|---|
| `String`, `Long`, … | the bytecode named it; the value was read **as** that type |
| `?opNNN` | an instruction touched the slot and is **not in our type table** — a gap, and `NNN` is the entry that would fix it |
| `?none` | an instruction touched it that provably conveys no type |
| `?unseen` | **no instruction touched the slot at all** |

**Two symbols, one meaning each.** `?` always means *this slot's type is unknown*. `~`,
closing the signature in `typetext`, always means *the walk did not read the whole body*:

| `typetext` | meaning |
|---|---|
| `Long,String` | the walk read the whole body |
| `Long,String~` | the walk **stopped early or resynchronised** past a statement |
| `Long,...` | the signature was too long for its buffer and was cut — fewer names than parameters |

The two compose. `Long,?unseen` says that parameter is genuinely never read.
`Long,?unseen~` says the walk was incomplete, so the same `?unseen` may instead be a load
the walk **skipped** — and nothing in the trace can tell those apart, which is exactly why
the `~` is there rather than a quiet guess. A `~` weakens every `?` in that row. It is rare:
across a full sweep of 5,266 VBA entry rows carrying a signature, not one walk was partial.

**String escaping, in `args` and `ret`, on XLL and VBA rows alike.** A decoded string is rendered
between quotes, with everything escaped that would otherwise be ambiguous, so
the value round-trips out of its CSV field unchanged:

| in the string | rendered as |
|---|---|
| `"` | `\"` |
| `\` | `\\` |
| tab, CR, LF | `\t`, `\r`, `\n` |
| any other character below 32, and DEL | `\xNN` |
| any character above 126 | `\uNNNN` |

The closing quote is therefore the only unescaped `"` in the value. Two of these
exist because the alternatives were silently lossy: a raw CR or LF would be
turned into a *space* by the CSV writer, claiming the string held a space where
it held a line break; and a character outside ASCII used to become `?`, which no
reader could tell from a question mark the string really contained. `\uNNNN`
keeps every value ASCII, so it reads the same in any viewer. A VBA string longer
than 256 characters is cut there, and ends with `...` after the closing quote.

**`parent` and `depth` — where a call sits in the chain.** Both are columns of
their own, beside `span`, because `parent` *is* a span: the three call-tree
fields read together. Both are on every row, entry and exit, on both sources.

- **`depth`** is how far down the call chain this activation was. `depth=1`
  means no other frame of the same source was open on that thread — this call
  started the chain. (Excel may well have been busy; the depth counts frames of
  one source, not everything in flight.) **Each source counts its own:** an XLL
  function a VBA procedure calls reads `depth=1` and `parent=0`, and its place
  under the VBA frame shows only in the order of the rows.
- **`parent`** is the `span` of the activation that called it, `0` at the top.
  Because `span` pairs an entry row with its exit, `parent=2` means *called by
  the activation whose two rows both say `span=2`*.
- **A macro Excel runs while VBA waits in `DoEvents`** — an `Application.OnTime`
  macro, an event — starts a chain of its own: `depth=1`, `parent=0`. It runs on
  top of the frame that called `DoEvents`, but that frame did not call it. The
  disarm log counts these as `doEventsChains`.

Together they place a row exactly:

```
seq  kind   span  parent  depth  function  ticks  trust
1    entry  1     0       1      Go
2    entry  2     1       2      Middle
3    entry  3     2       3      Leaf
4    exit   3     2       3      Leaf      412    exit
5    exit   2     1       2      Middle    980    exit
6    exit   1     0       1      Go        1533   exit
```

**Why both, when one nearly implies the other.** `depth` tells you the *shape*
of the chain; `parent` tells you *who*.

You could almost manage on `depth` alone: reading one thread's rows in `seq`
order, a `depth=3` entry belongs to the last `depth=2` still open. That
reconstruction is right until a row goes missing — and rows can go missing,
which is exactly what the holes in `input` record. After a gap, every later row
is quietly attached to the wrong caller. `parent` names the caller outright, so
the same missing row shows up instead as a `parent` pointing at a `span` that is
not in the file: a question, rather than a confident wrong answer.

`parent` also settles **recursion**, which `depth` cannot. Three activations of
one procedure produce three rows with the same `function` and the same `proc`;
only `parent` says which of them called which.

Going the other way, `parent` alone would turn "how deep did this get" into a
link-by-link walk of the whole file. `depth` is a property of the row itself, so
it survives a gap. There is no depth limit: each thread's frame stack grows as
the calls nest.

**`trust` — whether `ticks` is a measurement or an upper bound.** It names
**what ended the measurement**, not a verdict on it, so the cause survives and
the verdict is one lookup away. A column of its own, beside the number it
qualifies, now that `note` is gone.

| value | what ended the measurement | `ticks` is |
|---|---|---|
| `exit` | the exit opcode (VBA), or the return detour (XLL) | a measurement |
| `end` | the `End` opcode, which killed the whole chain | a measurement |
| `backstop` | the stack pointer, at the *next* statement | an **upper bound** |
| `flush` | the flush at disarm | an **upper bound** |
| `async` | nothing yet — the call was dispatched and has not answered | **absent**: `ticks` is empty |

`end` is grouped with `exit` and not with the other two abrupt closes on
purpose: the `End` opcode fires at the moment the session dies, so the frame is
closed when it actually ended rather than whenever the next statement happened
to arrive. It always accompanies `outcome` `abandoned`.

**Asynchronous XLL functions are NOT SUPPORTED.** An add-in function registered
asynchronously — a leading `>` in its type text, or an `X` (async handle)
argument — is recognised and its exit row is marked `trust` `async` with an
empty `ticks`, and that is the whole of it. **Do not rely on a duration, or on
any pairing with the answer, for such a call.**

The reason is structural rather than an omission we intend to close. An async
function returns *before its answer exists*: it hands Excel a handle, and the
result arrives later through `xlAsyncReturn`, on whatever thread the add-in
chooses. The `span` therefore measures the **dispatch**, not the work, and there
is no second event this tracer hooks that would let it close the pair. Reporting
the dispatch time would say a four-second call took microseconds, which for a
tool that exists to find slow things is worse than saying nothing — so `ticks`
is left empty and `trust` says which question it is not answering.

**This path is also untested.** No add-in in this repository registers an
asynchronous function, so the behaviour above is reasoned from the registration
flags, not measured. Treat an `async` row as a marker that a call happened, and
nothing more.

Real-time data (`=RTD(...)`) is a **different mechanism** — a COM server, not
the XLL C API — and is not hooked at all. It produces no rows.

Only the exit opcode fires when an activation actually ends. A fully unhandled
unwind fires none at all, so its frames are closed later and `ticks` runs until
whatever happened next -- measured at 9,164,013 ticks against 383 for a
comparable clean call. `qpc` is late on those rows for the same reason, which
matters because it is what orders the merged XLL/VBA timeline. The number is
kept rather than blanked, because it is a true bound; what would be wrong is
presenting it as a reading.

**Nesting.** Both sources state it outright, in `parent` and `depth`, each counting
its own frames per thread. The same tree can be read independently from the
ENTRY/EXIT INTERLEAVING -- span B is inside span A when A opens before B and
closes after it on the same thread, ordered by `seq`. So an add-in that
re-enters Excel via `xlUDF` reads as nested intervals, at `depth` 2 and 3 with
each `parent` naming its caller, while cell-level nesting (`=TxB(TxB(1,2),3)`)
reads as two disjoint calls at `depth` 1, because Excel evaluates the inner call
and finishes it before the outer begins. The one exception is a macro run inside
`DoEvents`: its rows fall inside the waiting frame's interval but begin a chain of
their own. `Get-MaxNestDepth` and the fuzz suite's
`NestDepth` check the columns against the interleaving -- a flattened tree would
pass entry/exit pairing, span uniqueness and the caller invariants unchanged.

## Values: arrays, Variants and types

One grammar spells every value in `args` and `ret`, on both sources, and one code path writes
it (`core::ValueWriter`), so the same value reads the same wherever it lands.

**An array is its element type, both bounds of every dimension, then one brace level per
dimension**, the first dimension outermost. For a 2-D array that means rows: the whole of row 1,
then row 2.

```
Long[0..3]{1,2,3,4}                                   1-D
Variant[1..2,1..3]{{11,12,13},{21,22,23}}             2-D: a range, rows first
Variant[1..1,1..3]{{1,2,3}}                           one row of a range is still 2-D
Long[0..1,0..1,0..2]{{{1,2,3},{4,5,6}},{{7,8,9},{10,11,12}}}   3-D
Variant[0..-1]{}                                      Array(): bounds 0..-1, no elements
Long()                                                a dynamic array never allocated: its type, no bounds
?[0..-1]{}                                            an empty ParamArray: its descriptor names no element type
```

- **Both bounds, always**, because `Option Base` decides what a bare count would mean:
  `Dim a(3)` is 0..3 under base 0 and 1..3 under base 1.
- **A bare brace is always a dimension; a nested array always carries its own header.** So
  `{{1,2}}` is one row of a 2-D array, and `{Long[0..1]{1,2}}` is a 1-D array holding an array.
  Nesting goes 32 levels deep, then reads `[...]`.
- **Every element is written.** Arrays are not truncated. The one limit is 4 MB for a single
  value: an array past it keeps its header and has no braces — `Double[0..99999999]` — so its
  shape is known and its contents are not in the file. Never part of the contents. The disarm
  log counts these (`values: N array(s) over the 4 MB value limit`).
- **An element that would not decode is `?`**, not a refusal of the array: the shape came from
  the descriptor and stays true whatever one element holds.
- **An array with no storage is still a value.** `Dim a() As Long` before `ReDim` reads `Long()`
  in a Variant, and `?()` as a `Function … As Long()` result, whose element type nothing records.
  In a `Ref&` argument a zero reads `0x0`: a `ByRef LongLong` holding 0 and an unallocated
  `ByRef` array are the same eight bytes.

**Inside a Variant, a bare number means Double; every other type is named.** This covers a
Variant argument, a Variant result, each element of a Variant array, and a Range's contents:

```
1.5                   Double
Integer(1)            Integer, and Boolean passed as an Integer
Long(7)  Single(1.5)  Byte(3)  LongLong(9)  Currency(1.5000)  Decimal(1.5)  Date(46352)
"x"   TRUE   #N/A   Empty   Null   Nothing                    already say what they are
Variant[0..3]{1234.5,"two",Long(3),TRUE}
```

A typed array's elements are bare, because the header already names them — `Long[1..3]{3,6,9}`
— and so is a declared scalar argument, `a1:Long=5`. `Range.Value2` hands back only Doubles,
strings, Booleans, errors and Empty, so a range's contents are never tagged.

An object reads `Class@0x…(where)`, and a Range, Collection or Dictionary whose contents were
read is followed by `=` and its value (see *Objects*). A user-defined Type reads `udt@0x…`. An XLL reference reads
`SRef(R2C2:R3C3)` or `Ref(R2C2:R3C3,R5C5:R5C5)`: its cells are not read, because `xlCoerce` on a
cell not yet calculated makes Excel abandon the call and run it again later.

## JSON Lines

With `FORMAT=JSONL` the trace file is `XRayXL_Trace_<id>_<pid>.jsonl`: no header, and one JSON
object a line whose keys are the CSV columns in the same order. `seq`, `span`, `parent`,
`depth`, `thread`, `qpc`, `argcount` and `ticks` are numbers; the rest are strings; an empty
field has no key. `args` and `ret` are structured, and **every value names its type**, a
Double included:

```
{"seq":12,"input":12,"kind":"exit","source":"VBA",...,"ret":{"t":"Array","elem":"Variant",
 "bounds":[[0,2]],"v":[{"t":"Long","v":7},{"t":"Double","v":2.5},{"t":"String","v":"x"}]},
 "rettype":"Variant","outcome":"returned","ticks":604,"trust":"exit"}
```

| value | JSON |
|---|---|
| a number | `{"t":"Double","v":1.5}`, `{"t":"Long","v":7}`. `Currency`, `Decimal`, `LongLong` and `UInt64` carry their exact digits as a string: `{"t":"Currency","v":"1.5000"}`. `NaN` and the infinities are strings too |
| a string | `{"t":"String","v":"x"}`, with `"cut":true` when the source was longer than was read |
| Boolean, error | `{"t":"Boolean","v":true}`, `{"t":"Error","v":"#N/A"}` |
| a word | `{"t":"Empty"}`, and the same for `Null`, `Missing`, `Nothing`, `AsyncHandle` |
| not known | `{"t":"Unknown","v":"?vt17"}` |
| an array | `{"t":"Array","elem":"Variant","bounds":[[1,2],[1,3]],"v":[[…],[…]]}`: `v` nests one JSON array per dimension, rows first. In a typed array (`elem` other than `Variant`) the elements are bare JSON values, since `elem` names their type; a cut string there is still an object. Over the value limit, `v` is replaced by `"omitted":"over the value limit"`. An element type the descriptor does not name is `"elem":null`, and an array never allocated has `"bounds":[]` and no `v` |
| an object | `{"t":"Object","class":"Range","ptr":"0x1E2…","where":"'[Book1]Sheet1'!A1:C2","value":{…}}` |
| a reference | `{"t":"SRef","areas":[[2,2,3,3]]}` |
| a UDT | `{"t":"Udt","ptr":"0x…"}` |

On an event row, each object names its parameter instead of a slot:
`{"name":"Target","type":"Range","value":{…}}`. The `arm` and `disarm` rows' values have no
`"type"` key; each value names its own `t` as always.

`args` is an array with one object per slot — `{"slot":1,"type":"Variant","value":{…}}`, with
`"address":"0x…"` after the type where the text has `@0x…`, and
`"unreadable":true` in place of a value that could not be read — and, when the XLL plan
describes fewer slots than exist, a closing `{"described":2,"slots":5}`.

## Excel errors

An error travelling through VBA — in an argument, in a return, or as an element
of an array — reads as the text the user sees in the cell:

```
args   a1:Variant=#N/A
ret    #DIV/0!
ret    Variant[0..2]{Integer(1),#VALUE!,Integer(3)}
```

`#NULL!`, `#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A` and
`#GETTING_DATA`. **Unquoted**, so it cannot be mistaken for the string `"#N/A"`,
which renders with its quotes. The same table names them on the XLL side, so an
error an add-in returned and an error a UDF was handed read identically.

**An SCODE outside that set keeps its number** — `Error(0x…)`. A COM error is
not a cell error and naming it would claim it was.

**`Missing` is not an error, though it arrives as one.** An omitted `Optional`
is a VARIANT carrying `VT_ERROR` exactly like `#N/A` does; the two are told apart
by the SCODE alone — `DISP_E_PARAMNOTFOUND` is facility `0x2`, an Excel error is
facility `0xA`. So a supplied `#N/A` can never be reported as an absent
argument, and an absent argument can never be reported as a cell error.

**A cell REFERENCE in a `Variant` parameter is an object, not a value.** Excel
hands a `Range` for `=MyUdf(A1)` and the value for `=MyUdf(NA())`, so the first
reads `object@0x…` and the second `#N/A`. VBA's own `IsError` and `CStr` coerce
through the Range's default property and will say "error" for both — they are
answering about the coerced value, not about what the parameter holds. The trace
reports what the parameter holds.

## Objects

With `OBJECTS` on — the default — an object argument or return is named, and a
`Range`, `Worksheet`, `Workbook`, `Collection` or `Scripting.Dictionary` is
described. The address is always kept, so one object can be followed from row to
row:

```
a1:Variant=Range@0x000001E2…('[Book1]Sheet1'!A1:C2)=Variant[1..2,1..3]{{11,12,13},{21,22,23}}
a1:Variant=Range@0x000001E2…('[Book1]Sheet1'!A:A)   -- addressed, deliberately not read
a1:Variant=Worksheet@0x000001E2…([Book1]Sheet1)
a1:Variant=Workbook@0x000001E2…([Book1])
a1:Variant=Collection@0x000001E2…=Variant[1..3]{1.5,"x",Nothing}
a1:Variant=Dictionary@0x000001E2…=Variant[0..1,0..1]{{"a",Integer(1)},{Long(2),Empty}}
a1:Variant=Position@0x000001E2…                     -- named; no detail we know how to fetch
a1:Variant=object@0x000001E2…                       -- OBJECTS off, or nothing worked out
a1:Object=Nothing                                   -- no object at all
```

**This is the one setting that makes the tracer TALK to Excel.** Everything else
reads memory or asks the flat C API, which is passive. This calls the object
model — `Address`, `Count`, `Value2`, `Name`, and a container's `Keys`, `Items`
or enumerator — on the calculating thread, at a statement boundary. `XRayXL_SetTraceParam VBA, "OBJECTS", FALSE` returns the
tracer to pure observation, and every object then reads `object@0x…`.

**A class is identified, never guessed.** `QueryInterface` against the published
interface ids — Excel's, VBA's `_Collection` and the Scripting Runtime's
`IDictionary` — decides whether a detail may be fetched — not the presence of an `Address` property, and not a familiar-looking
vtable. An interface id that is wrong, or that Excel changes, simply never
matches and the object falls through to being named.

**The name is the CLASS, from `IProvideClassInfo`** — the same question VBA's
`TypeName()` asks. `IDispatch::GetTypeInfo` answers with the default *interface*
instead, which is `_Collection` where VBA says `Collection`; trimming that
underscore would be a guess about a naming convention, so the class comes from
the source that actually knows it. Where no class info exists, the interface name
is reported as it comes, underscore and all.

**A range's cell count is checked BEFORE its value is asked for.** `A:A` is over
a million cells and reading it would materialise a ~25 MB array inside a
calculation. Over the ceiling of 4,096 cells — or when the count
could not be had at all — the address is reported and the contents are not, which
is visible in the row: no `=` follows the address. The same goes for a range of
several areas (`A1:A2,C1:C2`), whose `Value2` is the first area's alone.

**What Excel hands back for `Value2` is measured, not assumed.** A single cell
gives a **scalar**. Every multi-cell range gives a **2-D** array, including a
single row — `A1:C1` is `[1..1,1..3]{{1,2,3}}`, *not* a 1-D array of 3. Rows come
first, one brace level each: `A1:C2` reads `{{11,12,13},{21,22,23}}`.

**A Collection reads as a 1-D Variant array from 1**, in the order `For Each`
gives. Its keys are not shown: a Collection has no way to give them back. **A Dictionary reads as a 2-D Variant
array of `{key,item}` rows**, in `Keys` order, the rows numbered as `Keys()`
numbers them. Every key and item is read as a Variant array's element is, so each
is any type — a number named, a string quoted, `Empty`, `Nothing`, an array, or
another object described by these same rules. An empty one keeps its shape:
`Variant[1..0]{}` and `Variant[0..-1,0..1]{}`.

**A container's contents are declined, not cut**, as a range's are: over 4,096
items, or when its items cannot all be read, it is named with no `=`. A container
already being read — one that holds itself, or holds one that holds it — is named
where it recurs rather than read again, and past eight containers deep the next is
named only.

**Any failure renders the address.** A call that fails, a class with no name, the
setting off — all produce `object@0x…`, which is what this always said and is
never wrong.

**Addresses here are Excel's own**, from `Range.Address(…, External:=True)` —
and `callerref` now matches them, quoting included. See *Quoting* below.

## Quoting an external address

An address — in `callerref`, and in a described `Range` — is quoted exactly where
Excel quotes it, so it can be pasted straight back into a formula. Measured
against Excel rather than recalled:

| | quoted |
|---|---|
| `[Plain1.xlsx]Sheet1!A1` | no — **a dot alone does not quote**, so an ordinary saved workbook is bare |
| `[Under_score.xlsx]Under_1!A1` | no — underscore is safe |
| `[Plain5.xlsx]A.B!A1` | no — a dot in the *sheet* is safe too |
| `'[has-hyphen.xlsx]Sheet1'!A1` | yes — a hyphen, either side |
| `'[has space.xlsx]Sheet1'!A1` | yes — a space, either side |
| `'[Digits123.xlsx]1Sheet'!A1` | yes — a **sheet name starting with a digit** |
| `'[Plain4.xlsx]Bob''s'!A1` | yes, and the apostrophe is **doubled** inside |

So the safe set is letters, digits, `_` and `.`; anything else in either name
quotes, as does a sheet name beginning with a digit.

**The doubt falls on the side of the quote.** Quoting where Excel would not is a
cosmetic difference in a reference that still pastes back; failing to quote where
Excel would produces one that does not.

**A reader needs to accept both forms.** An address has exactly one `!`,
separating the prefix from the reference, and neither half can contain another —
so splitting on it works whether or not the prefix is quoted.

## Reading `args`, and what `?` means where

**In CSV, `args` is a small format inside one field.** (JSON Lines writes it as
an array of objects instead; see *JSON Lines*.) Values are
`a<N>:<type>=<value>`, joined by single spaces, where `<N>` is the SLOT and not
the parameter ordinal — a `ByVal Variant` is a 24-byte VARIANT across three
slots, so the parameter after one is `a4`.

**`@0x<address>` between the type and the `=` says where the argument's storage
is**, on VBA rows only:

| the slot | the address |
|---|---|
| a `ByRef` parameter, `a1:Long&@0x…` | the pointer the slot holds: the caller's variable |
| `?none` that the body passes on `ByRef` as it is (opcode 751) | the pointer the slot holds |
| `?none` that the body passes on `ByRef` from its own copy (opcode 671) | the slot's own address |

It is there so a parameter the body only passes on can be typed from the
procedure it was passed to. **The same address on a row and on a row beneath it
in the same call tree is the same storage**, and a `ByRef` argument must match
its parameter's type exactly, so the callee's type is the parameter's. A
`?none` with a 671 address is a `ByVal` parameter; with a 751 address, a `ByRef`
one. Match only within the parameter's own activation — rows whose `parent`
chain leads to its span — because an address is reused once a frame has ended.

```
span 9   U02  a1:?none@0x26E85895F90=0x5        ByVal, passed on from its own copy
span 10  CL   a1:Long&@0x26E85895F90=5          parent 9: the same storage, so U02's a1 is a ByVal Long, 5
```

Only a quoted string can contain a space, and a string's own quotes and
backslashes are escaped, so **splitting on ` a<N>:` outside quotes is
unambiguous**. Track quote depth; do not split on spaces alone.

**`?` appears in several places and means something different in each.** It is
always "this is not known", and the position says what *this* is:

| where | what is not known |
|---|---|
| the type position — `a1:?unseen=…` | the parameter's TYPE. The word says which not-knowing (see above) |
| before an array's bounds — `?[0..2]{…}` | the array's ELEMENT TYPE could not be named |
| inside an array — `{1,?,3}` | that one ELEMENT would not decode. Deliberately not a refusal of the whole array: the shape and size came from the descriptor and stay true whatever one element holds |
| a typed array's return — `?()` | the element type of an `Elem()` return could not be named |

On the **XLL** side a type position is the registration's own code (`Q`, `B`,
`D%`…), which is authoritative — so a `?` never appears there in a traced row.
`<N>` counts slots there too: an `O` array takes three, so the argument after one
is `a4`. An XLL array reads like a VBA one, rows first with both bounds:
`Double[1..2,1..2]{{1,2},{3,4}}` for an `FP` or `O` argument, `Variant[1..R,1..C]{{…},…}`
for an `XLOPER` array.

## What the trace cannot tell you on its own

**Whether argument capture was OFF, or tried and declined.** On VBA rows both
leave `typetext`, `argcount` and `args` empty, and the trace file carries no
record of the setting. XLL rows keep `typetext` and `argcount` with `ARGS` off:
they come from the registration, not from capture. The **log** does — `trace param set: VBA ARGS` at arm,
and `VBA args: N captured` with the decline reasons at disarm. If a whole
session has no arguments anywhere, read the log before concluding anything
failed.

This is the one place the three channels are not independent: the trace is the
data, and for this question the log is the only witness.

**`argcount` distinguishes what it can.** `0` means the procedure takes no
parameters — a fact — and empty means the capture did not run or declined. The
signature does the same: `()` is "no parameters", empty is "not recovered".

## Limitations

**The declared TYPE of an object parameter is always just `Object`.** VBA's
bytecode selects the load instruction by *kind* — "an object reference" — not by
class, so `As Range`, `As Worksheet` and your own class modules all read
`Object` in `typetext`. The class comes from the VALUE instead, not the
declaration — see *Objects* below.

**A UDT is reported as the record's address, with no members.**

```
typetext   (Udt&)
args       a1:Udt&=udt@0x000001E226D99100
```

A user-defined `Type` is always passed by reference and the slot points straight
at the record. Its fields are not walked, so an object field — a `Range` inside a
`Type` — does not appear at all.

## Ordering guarantees

- `seq` is strictly increasing and equals file order.
- Sorting by `qpc` is sound because — asserted, not assumed — within a thread
  `qpc` never runs backwards, and within a span the entry is stamped before
  its exit (and numbered before it, too).
- Across threads under multithreaded calc, `seq` and `qpc` genuinely disagree
  (measured: 213 inversions in 809 row boundaries). That is the design, not a
  defect: `seq` orders the file, `qpc` orders the events.

## Asserting against volatile columns

Stable across runs: `kind`, `module`, `function`, `typetext`, `caller`,
`argcount`, the shape of `args`, `ret`, `rettype`. Volatile by nature: `seq`,
`span`, `thread`, `qpc`, tick counts, trailer addresses, pid-stamped workbook
names. Tests assert *values* on the stable set and *relations* on the
volatile set — pairing, ordering, reconciliation against the totals line.
This is why there is no golden file.

**A row must carry every column.** The reader counts the fields on each line
and refuses the file if any row disagrees with the header — a row short by one
would otherwise be padded with blanks and every value after the gap silently
read as the column before it.

The row invariants hold for *any* correct trace and are asserted on
every row by every driver, via `Test-RowInvariants` in
`tests/sweep/_xray_common.ps1`: an entry-kind row always names its caller; the
kind is one of the closed set above; a `callerref` is present exactly when
the kind says one should be; a `cell` description is a full external
address and not a bare reference; an exit row carries neither. Drivers
additionally assert the *expected* caller of the first traced frame, keyed
on how the test triggered it (a formula: `cell`; Run/Evaluate and event
handlers: `none`).

The same function holds `outcome` to its closed set: every exit row carries
one, an XLL exit row carries only `returned`, and no other row carries any — so
a `threw` on an XLL row, where no error chain is followed, is a test
failure rather than a plausible-looking value.

**`caller` answers "what started this chain", not "what called this
frame".** Excel *inherits* it: a `Worksheet_Change` that fires because a
button's macro wrote a cell reports the **button**, not its own absence of
a caller — measured for a graphic object, a toolbar button and a
right-click item alike, with a `Worksheet_Change` that had no outer macro
correctly reporting `none`/`ref` as the control. So an event cascade names
whatever started it, all the way down, which is usually the question worth
asking. To tell a root from an inherited frame, read `depth`: `1` is the
root. On the XLL side `depth` counts XLL frames on the thread, and a thread
nested past 64 XLL calls writes no rows for the deeper ones, with no marker.

## Evolution policy

The header line **is** the version. Any change — a column (append-only, at
the end), a new kind, a new format — lands as one change touching four
places: `kHeader` in `rowcsv.cpp`, this document, `Read-TraceFile`, and
`tests/sweep/format/reader-contract.test.ps1`. An old reader meeting a new file
fails the header check and *says so*; nothing skips silently. A future format
(a richer JSON, should one land) adds a dispatch
branch to the reader and a section here; the **model** — the kinds, the
guarantees, the channel split — is format-independent.

## What buffering guarantees about the file

The settings themselves are in [TraceOptions.md](./TraceOptions.md). What
matters *to a reader of the file* is what each choice guarantees:

| | The file is | Loss is |
|---|---|---|
| `BUFFERSIZE=0` | complete, and written as it happens — a crash mid-calc leaves everything up to it on disk | impossible |
| `BUFFERSIZE=N`, `BUFFERWHENFULL=PAUSE` (the default) | complete | impossible; a full ring makes the calc wait until it is half empty |
| `BUFFERSIZE=N`, `BUFFERWHENFULL=DROP` | possibly short | **out of band** — the count from `XRayXL_Disarm`, the status summary and the disarm line; located in the file by holes in `input` |

Under `DROP` the reconciliation invariant still holds exactly: rows written
plus rows dropped equals `framesOpened + framesClosed`. A dropped entry does not
take its exit with it, so an exit row can appear with no entry in the file. **The file itself
carries no marker row** — a phantom row would be an event that never happened.

**Disarm drains.** Read-after-disarm completeness survives buffering, whatever
the setting. That is the fence every suite stands behind, and it is why the
suites run the shipped configuration rather than a special unbuffered test mode.
