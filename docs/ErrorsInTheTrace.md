# How errors read in a trace

Every VBA call in a trace ends with an `exit` row, and that row's `outcome` column says
how the call ended. When nothing goes wrong it says `returned`. This page is about the
rest: what each word means, how a chain of calls reads when an error passes through it,
and the cases where VBA or Excel handle an error in a way you might not expect.

The column reference is in [TraceRowModel.md](./TraceRowModel.md); how the tracer works
out an outcome is in [VBATracing.md](./VBATracing.md).

---

## The six outcomes

| `outcome` | what happened to the call |
|---|---|
| `returned` | it finished normally |
| `threw` | an error was raised in it, and it went no further |
| `unwound` | an error raised below it passed through it on the way out; it ran nothing after that |
| `handled` | it caught an error raised below it, and carried on |
| `unhandled` | an error left VBA through it into a worksheet cell, which shows `#VALUE!` |
| `abandoned` | VBA stopped it: it neither returned nor passed an error on |

An XLL function's exit row always reads `returned`, because it is written only when
the function returns.

**Read `outcome` together with `trust`.** `trust` says what closed the call, and so
whether its `ticks` is a measurement or only a ceiling:

| `trust` | the call was closed by | `ticks` is |
|---|---|---|
| `exit` | its own ending | a measurement |
| `end` | the `End` statement | a measurement |
| `backstop` | the next VBA statement to run, after the call had already gone | an upper bound |
| `flush` | disarming | an upper bound |

A call that ended in an error runs no ending of its own, so it is always closed by
something else: `backstop`, `flush` or `end`.

---

## An error caught by a caller

The common case. `Report` has an error handler and calls `Load`, which calls `Parse`,
which raises:

```vba
Sub Report()
    On Error GoTo Failed
    Load
    Exit Sub
Failed:
    Debug.Print Err.Description
End Sub

Sub Load()
    Parse
End Sub

Sub Parse()
    Err.Raise 5
End Sub
```

| function | depth | outcome | trust |
|---|---|---|---|
| `Parse` | 3 | `threw` | `backstop` |
| `Load` | 2 | `unwound` | `backstop` |
| `Report` | 1 | `handled` | `exit` |

Read it outwards from the raise: `Parse` raised it, it passed through `Load`, and
`Report` caught it. `Parse` and `Load` never ran their endings, so they are closed when
`Report` runs its next statement.

**An error caught in the same procedure reads `returned`.** With `On Error Resume Next`,
or a handler in the procedure that raised, nothing leaves the call, and it does return.
`threw`, `unwound` and `handled` describe an error crossing from one call to another.

**A handler that raises again.** If `Report`'s handler raised a new error that left it,
`Report` reads `threw`, and whichever call catches that one reads `handled`.

---

## An error into a worksheet cell

A VBA function used in a cell formula has no VBA caller to pass an error to. Excel
catches it and the cell shows `#VALUE!`:

| function | depth | outcome | trust |
|---|---|---|---|
| `Parse` | 2 | `threw` | `backstop` |
| `PriceOf` | 1 | `unhandled` | `backstop` |

`PriceOf` is the function the cell called, so it reads `unhandled`: the error left VBA
there. The functions it called keep `threw` and `unwound`, so the raise is still found.
The calls close when Excel starts its next VBA call, or at disarm (`trust=flush`).

- **A macro that recalculates the cell reads `returned`.** The error stopped at the cell;
  the macro never saw it.
- **A function that returns an error value raised nothing.** `PriceOf = CVErr(xlErrValue)`
  reads `returned`, with `#VALUE!` in the `ret` column.
- **`Application.Evaluate`, defined names and conditional formatting count as cells.** A
  function run through any of them is given a calling cell, so its error reads
  `unhandled`. Under `Evaluate` that cell is the active cell, not one holding a formula.

---

## An error nothing catches: VBA's error dialog

An error that no call catches, in a macro, a sheet event, an `Application.OnTime` macro or
a macro started with `Application.Run`, makes VBA stop and show its *Run-time error*
dialog. **Pressing End** stops all VBA at once. No call runs its ending, and the tracer
sees nothing happen at that moment. The calls stay open until the tracer finds them gone.
That happens when the next VBA call starts, or at disarm, whichever comes first. Either
way they read the same:

| function | depth | outcome | trust |
|---|---|---|---|
| `Parse` | 3 | `threw` | `backstop` or `flush` |
| `Load` | 2 | `abandoned` | `backstop` or `flush` |
| `Report` | 1 | `abandoned` | `backstop` or `flush` |

`Parse` raised the error. `Load` and `Report` read `abandoned`: the error never passed
through them, and End stopped them where they were. Their `ticks` run until the calls
were found closed, which includes however long the dialog was open.

A call chain that Excel started for a cell is the exception. Its error went to the cell,
so it reads `threw` … `unhandled` as above, and only the calls beneath it read `abandoned`.

**Pressing Debug** stops VBA in the editor instead, with every call still running:

- **Continue** (F5 or F8), and the calls carry on and close normally.
- **Reset**, and they read as if you had pressed End.
- **Disarm while VBA is stopped there**, and they are closed as still running (`returned`,
  `trust=flush`).

These need a person at the editor, so no automated test covers them.

### Errors that do not reach their caller

Two situations look as if an error should go back to the caller's `On Error`, but VBA
stops at the dialog instead. The caller reads `abandoned`, even with a handler:

- **A macro called with `Application.Run`.** An error in the called macro, with no handler
  of its own, shows the dialog in that macro. The caller's `On Error Resume Next` never
  gets it. A direct call (`Load` rather than `Application.Run "Load"`) does return the
  error to the caller.
- **A sheet event fired by a macro.** When a macro writes a cell and the sheet's
  `Worksheet_Change` raises an error, the dialog appears in the event handler. The macro
  that wrote the cell is stopped too.

**An error in a call Excel started on its own reads `threw`.** A sheet event fired by a
user's edit, or an `OnTime` macro, is the first call in its chain, so the raiser is also
the outermost call. It reads `threw`, not `unhandled`: only a cell formula makes an error
`unhandled`.

---

## The End statement

`End` in your code stops all VBA, as the dialog's End button does. Here the tracer sees it
happen, so every open call closes at once, `abandoned` with `trust=end`. That includes the
call that ran `End`: no error was raised, so no call reads `threw`.

---

## Stopping VBA from the editor

- **Reset after a breakpoint or a `Stop` statement** reads like End, with one difference:
  the call that was stopped reads `abandoned` rather than `threw`, because it raised
  nothing. The tracer knows it stopped from the `breaks` count, which is kept whether or
  not the `breaks` column is switched on (`BREAKPOINTS` in
  [TraceOptions.md](./TraceOptions.md)).
- **Time spent stopped in the editor is part of the call's `ticks`.** Switch the `breaks`
  column on to see which calls include it.

These cases need a person at the editor, so no automated test covers them.

---

## Edge cases and limits

- **Ctrl+Break.** A user break is not forced. If the code sets
  `Application.EnableCancelKey = xlErrorHandler`, it arrives as error 18, which reads like
  any other error. Otherwise VBA shows *Code execution has been interrupted*, and pressing
  End there should read as End on the error dialog, which means the call that was running
  reads `threw` though nothing raised an error: the tracer cannot tell the two dialogs apart.
  This one needs a person at the keyboard, so it has not been measured.
- **Disarming while a macro runs.** A macro still running at disarm, for example one that
  disarms itself, or one waiting in `DoEvents`, reads `returned` with `trust=flush`. It
  never reads `threw`.
- **Excel calling a function twice.** When a formula needs a cell that is not yet
  calculated, Excel lets the function finish and calls it again later. That is two whole
  `returned` calls, not an abandoned one.

---

## How the tracer tells a stopped call from a running one

VBA gives no signal when its dialog's End button is pressed, so the tracer works it out
from the stack.

- **When the next VBA call starts**, and none of the open calls is still on the stack,
  VBA stopped them all. Only End (or Reset) stops every call at once; an error that is
  caught always leaves the catching call running.
- **At disarm**, the tracer walks the thread's real call stack. A call whose place on the
  stack now belongs to other code is gone. The walk only reads code that belongs to a
  loaded module, and stops at the first address that does not. Some calls sit beyond that
  point, as under `Application.Run`. For those, the tracer checks whether the call's own
  marker is still in place on the stack.

The details are in [Implementation.md](./Implementation.md), under *A stack that dies
whole was ended by VBA* and *A trailer can outlive its frame*.
