# How VBA tracing works

This explains how XRayXL traces VBA — what it changes in Excel, why, and how it
turns that into a row per call with arguments, return values and error outcomes. It
is written for a programmer who has not spent time inside interpreters or call
stacks. It builds up in layers: read the first section for the shape of the whole
thing, then go deeper only as far as you need. Register names and byte offsets are
kept out; where a hardware detail matters it is described in words.

For the exact opcode numbers, offsets and the evidence behind each claim, see
`docs/Implementation.md` (Part 4).

---

## The picture in one page

VBA code is not run as machine code. Excel compiles it to **p-code**, a compact
list of instructions for a small interpreter that lives in a Microsoft DLL
(`VBE7.DLL`). To run your macro, the interpreter walks that p-code one instruction
at a time. For each instruction it looks up a handler in a big **dispatch table** — an
array of function pointers, one per instruction kind — and jumps to it.

That table is the seam XRayXL works on. We do **not** rewrite any of Excel's or
VBA's code. We change a handful of pointers in the dispatch table so that, for a few
chosen instruction kinds, the interpreter first calls a small piece of ours and then
carries on into the original handler exactly as before. The instruction kinds we
intercept are the ones that mark meaningful moments:

- the **start of a statement** — so we can see each procedure begin and each line run;
- the **exit** of a procedure — so we can see it end and read its return value;
- the **End** of a VBA session.

From that stream of moments we rebuild, on the side, a picture of what VBA is doing:
which procedure is running, who called it, how deep the call nesting is, what
arguments came in, what value went out, and — when something goes wrong — where an
error was thrown and where it was caught or escaped. That side picture is kept in a
per-thread structure we call the **shadow stack**. Every finished call becomes one
entry row and one exit row in the trace file.

```
   VBA source  --compiled to-->  p-code (instructions for the interpreter)

   interpreter runs p-code:   for each instruction, look up its handler and jump

        dispatch table (function pointers, one per instruction kind)
        +-----------------------------------------------------------+
        |  ...  | start-of-statement | ... | exit | ... | End | ...  |
        +--------------^-------------------------^---------^--------+
                       |                         |         |
        we swapped these few pointers to point at our stubs first;
        each stub calls our recorder, then jumps to the original handler
```

---

## Layer 1 — What we patch, and why that way

**We swap pointers in a table; we never edit code.** Making the table writable,
storing a few new pointers, and putting the protection back is the whole change.
Nothing in Excel's or VBA's instruction stream is altered. This matters for three
reasons.

- **It is reversible and cheap.** Disarming puts the original pointers back. There
  is no patched machine code to unpick.
- **It survives Excel's own defences.** Modern Excel checks that indirect calls go
  to approved targets and keeps a protected copy of return addresses. Rewriting code
  would trip those and terminate the process. A pointer swap in a data table does
  not touch the instruction stream, so none of those defences fire.
- **It is one mechanism for everything.** We don't need to know anything about a
  particular macro. Every procedure in the process runs through the same interpreter
  and the same table, so patching the table traces all of them at once.

**What a patched slot does.** When the interpreter reaches an instruction whose
handler we replaced, it jumps into our small stub instead. The stub saves the state
the interpreter was using, calls our recorder (ordinary C++), and then jumps on into
the **original** handler as if nothing happened. The interpreter never notices; the
only cost is the detour.

**Which slots, and how many.** The table has about 1,700 slots. We patch a small
set of them, grouped by role: the two that begin a statement, the ones that end a
procedure (there are several exit forms), and the one that tears the session down
(`End`). Errors need no slot of their own: a call an error ends is the one that
never runs its exit. Everything the tracer knows is built from those
few moments.

**A note on trust.** We do not hardcode which slot is which. At arm time the tracer
*derives* the table's location and the meaning of the slots from the shape of the
loaded VBA DLL, and refuses to patch anything unless the structure checks out. If it
cannot be sure, it trace nothing rather than patch the wrong slot. That derivation is
its own topic; here it is enough to know the arming step is guarded.

---

## Layer 2 — The shadow stack, and why we need one

The obvious question is: if we see every statement, why keep our own stack? Why not
just note "we are now in procedure X"?

Because "which procedure" is not the same as "which **call** of that procedure".

**A procedure can be running several times at once.** A recursive function that
calls itself is the clearest case: the interpreter is inside `Factorial` five levels
deep, and all five levels are the same procedure. If we only recorded the procedure's
identity, those five nested calls would look like one, and we could never pair an
entry with its matching exit or measure how long the innermost call took.

**So each running call gets its own frame on a shadow stack.** A *frame* is a small
record: which procedure, when it started, how deep it is, which call it was made
from, and a few flags we will meet later. The shadow stack is just a per-thread array
of these frames — one thread's calls never mix with another's. It grows when a call
begins and shrinks when a call ends, mirroring what the interpreter is really doing.

```
   depth
     1   RunChain            <- a button macro, called by Excel
     2     OptionBook        <- called by RunChain
     3       RiskWeighted    <- called by OptionBook
     4         WeightFor     <- the call running right now (top of the stack)
```

**How a frame opens.** Every genuine call begins by running the procedure's first
statement, and we see that through a start-of-statement moment. When we notice a
statement belongs to a call we are not already tracking, we push a new frame. To tell
one *call* from another — the recursion problem above — we identify a call by the
combination of *which procedure* and *how far down the interpreter's own working
memory the call sits*. Two calls of the same procedure sit at different depths, so
they are distinct, and a genuinely recursive descent pushes a new frame each time.

**How a frame closes.** Most procedures end by running an exit instruction, which we
also intercept: that closes the matching frame and is the natural place to read the
return value. Some calls never reach a clean exit — an error can tear out of a
procedure without running its normal ending. For those we have a backstop: when we
later see activity at a shallower depth, we know the deeper calls must have finished,
and we close them. A final sweep at disarm closes anything still open.

**What the shadow stack buys us.** From it we get, for free, the call tree (who
called whom), the depth of each call, accurate timing for each call, and — crucially —
the ability to follow an error as it travels outward through the calls that were in
flight when it was raised. That last one is Layer 4.

---

## Layer 3 — Arguments and return values

**Arguments are read once, at the first statement of a call.** When a frame opens, the
call's arguments are sitting in the interpreter's working memory for that call, in a
known place. We read them there, before the procedure's body has had a chance to
overwrite them, and record them on the entry row. A by-reference argument (one the
procedure can change and hand back) is noted, so that at the exit we can re-read it
and report whether it actually changed.

**Return values are read at the exit, and the type is the hard part.** There is no
single place that always holds "the return value". Depending on the procedure, the
result might live in one slot for some kinds of function and a different slot for
others, and a plain `Sub` has no return at all. Reading a fixed slot blindly would
make a `Sub` appear to return a leftover number. So the tracer works out the *type*
first, and only reads a value if it can vouch for one:

- **The exit instruction often names the type.** Many typed functions leave through
  a type-specific exit, and that tells us what kind of value to expect. An exit that
  means "no result" (a `Sub`, a property assignment) makes us report nothing.
- **When the exit is generic, the instruction that stored the result names the
  type.** Class and form methods leave through a generic exit, so instead we look
  back for the instruction that wrote the result and read the type from it. If two
  candidate writes disagree, we refuse rather than guess.
- **A `Variant` is found by where it lives, not by how it was written.** A `Variant`
  result always sits in one particular place, and we read the kind tag it carries.

The upshot: the trace reports a return value only when the type is established, and
leaves the column empty otherwise. Reporting nothing is preferred over reporting a
confident wrong value.

**Returning an error value is not an error.** A function can deliberately return a
worksheet error, for example `CVErr(xlErrValue)`, which shows as `#VALUE!` in the
cell. That is an ordinary return: the function finished, and the trace records it as
`returned` with the error value in the return column. This is different from a
function that *fails*, which is Layer 4 — and telling the two apart is exactly why the
outcome column exists.

---

## Layer 4 — Errors: thrown, caught, and escaped

This is the subtle part, and worth reading slowly. The goal is that every finished
call carries an **outcome** that says how it ended, from a fixed list:

| outcome | meaning |
|---|---|
| `returned` | finished normally |
| `threw` | an error was raised in this call and it did not finish |
| `unwound` | an error passed through this call on its way out; it did nothing after the raise |
| `handled` | this call caught an error raised below it and carried on |
| `unhandled` | an error left VBA through this call into a worksheet cell (the cell shows `#VALUE!`) |
| `abandoned` | the `End` statement tore the session down; the call neither returned nor threw |

**Why the outcome is decided when the call closes.** An error can be raised by
`Err.Raise` or by VBA itself, in the middle of a division, a conversion or an
overflow, and there is no single place where every error starts. But every error
that leaves a call ends it the same way: without its normal ending. So the tracer
decides the outcome when the call closes:

- A call that ran its normal ending `returned`, even if it caught an error of its
  own on the way (`On Error Resume Next`).
- The innermost call that ended with no normal ending **threw**. `End`, and a
  macro still running when you disarm, are the two exceptions, and are told apart.
- The calls outside the thrower, still in flight, are **unwound** as the error
  passes through them — unless one of them runs again afterwards, which means it
  **caught** the error, and it reads `handled`.

Read outward from the raise, a chain looks like this:

```
   depth
     1   Caller        handled     <- ran its error handler; caught it
     2     Middle      unwound     <- the error passed straight through
     3       Thrower   threw       <- the error was raised here
```

This is where the shadow stack earns its keep: only by holding the calls that were in
flight can we follow an error outward and label each one.

**When nobody catches it: escaping into a cell.** The case that started this whole
design is a function called from a worksheet cell that fails with no handler. Excel
does not show an error dialog for that; it simply puts `#VALUE!` in the cell and moves
on. If a macro had triggered the recalculation, that macro keeps running and never
sees the error. So the trace must say two things: the function that failed reads
`unhandled` (it left VBA into the cell), and the macro that recalculated reads
`returned` (it was never affected).

```
   A button macro recalculates a sheet, and a cell's function then fails:

     RunChain        returned    <- a button started it; no calling cell
       OptionBook    unhandled   <- Excel computed a cell with it; it failed -> #VALUE!
```

On the shadow stack `OptionBook` sits *beneath* `RunChain`: the macro's call to
"recalculate" is still open when Excel computes the cell, so the failing function is
nested under the macro. That is why the simple idea — "an error that reaches the
bottom of the stack escaped" — does not work here: the error never reaches the
bottom, because the macro is under it. We need to recognise the failing call itself
as the escape point.

**How we know a call is "one Excel started for a cell".** This is the crux. An error
should stop and become `#VALUE!` only when the failing call was one Excel entered to
compute a worksheet cell — not when one VBA procedure simply called another. To tell
them apart the tracer asks Excel, for each call, **which cell it is computing** (a
documented lookup). A worksheet function and any helpers it calls are all computing
the same cell; a call whose cell is *different from the call beneath it* — or which
has a cell where the call beneath it has none, as with `OptionBook` under the
button-started `RunChain` above — marks the point where Excel handed control to VBA.
An error leaving that call becomes the cell's `#VALUE!`; an error inside a plain
VBA-to-VBA call keeps propagating normally.

That single question — "is this call's cell different from the one beneath it?" —
turned out to be the reliable signal after several tempting alternatives failed. The
alternatives and why they broke are recorded in `docs/Implementation.md`; the short
version is that trying to read the machine stack directly was fooled by leftover data,
and measuring how much stack a call used could not separate a class-method call from a
worksheet call. Because this check needs the calling cell on every call, resolving it
is not optional and cannot be turned off.

---

## Edge cases

The mechanism above is clean; reality has corners. These are the ones the tracer
handles deliberately, each covered by a test.

- **Recursion.** Handled by identifying a *call* rather than a *procedure*, so a
  function calling itself pushes a new frame each level.
- **Very deep nesting.** The shadow stack grows as calls nest, so every level of a
  deep recursion is recorded.
- **The `End` statement.** `End` stops all VBA immediately, running no endings at
  all. We intercept it directly and mark every still-open call `abandoned`, because
  they neither returned nor threw.
- **A same-procedure `On Error Resume Next` of a real error.** The call caught its
  own error and runs its normal ending, so it reads `returned`: nothing left it.
- **An error passed an Excel object.** Passing something like a `Range` to a VBA
  method used to confuse an earlier design; the calling-cell approach is immune,
  because the method is not computing a new cell.
- **`Application.Run` and object-model method calls.** These are VBA-to-VBA in
  effect: no new cell, so an error in them propagates to the caller rather than
  becoming `#VALUE!`.
- **A macro still running when tracing is disarmed.** Disarm closes open calls. A
  call that is genuinely still executing (a macro that called disarm, say) must not be
  mistaken for one that threw; the tracer checks whether the call is really still live
  before deciding.
- **A call left over from an earlier unhandled error.** The opposite of the above: a
  dead call that never ran its ending must be closed before the next call opens, so
  the next thing Excel computes does not appear nested under a corpse.

---

## The one known gap

An error becomes `unhandled` only when the failing call has a calling cell. Some
things Excel starts have **no** calling cell: a sheet event fired by a user edit
(`Worksheet_Change` and the like), and an `Application.OnTime` macro. An unhandled
error in one of those reads `threw` rather than `unhandled`. This is under-labelled,
not wrong — the failing call is still identified, the escape is still counted, and
Excel shows its own dialog. Closing this gap would need a different signal than the
calling cell, and none cheap and reliable is in hand. An event *triggered by VBA* (a
macro writes a cell, firing an event beneath it) is not affected: there the error
genuinely propagates to the macro, and the trace shows that.

---

## What VBA tracing deliberately does not do

- It does not modify your workbook or its code.
- It does not follow an error's *identity* — the outcome column says where an error
  was thrown and who caught it, never which error number it was.
- It does not dereference objects; an object argument is described by its class and,
  for a few known types, an identifying detail, but the tracer never calls into your
  objects to inspect them beyond that.
- It runs 64-bit Excel only.

See `docs/TraceOptions.md` for the settings, `docs/TraceRowModel.md` for the exact
shape of each row, and `docs/Implementation.md` for the mechanism in full detail.
