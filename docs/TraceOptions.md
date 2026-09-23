# Trace options — controlling what gets recorded

Everything XRayXL captures is set through registered functions, all called
through `Application.Run`; the setters refuse a cell. This document is the
reference for those settings and for the add-in's own log. What the resulting
rows *mean* is [TraceRowModel.md](./TraceRowModel.md).

The **XRayXL group** on the Developer tab is a front end to exactly these calls —
**Arm**, **Disarm** and **Options**, the last a dialog with five pages: Capture (a section
per source, each a Depth drop-down and check boxes), Output (the format
drop-down, the trace folder and file), Advanced (the output buffer, the optional `breaks` column and the log level),
About (the version and the licence), and Notices (the third-party notices). It holds no settings of its own, so the two can never
disagree: press Apply and `XRayXL_GetTraceParam` reports what you chose; change
something from a macro and the dialog shows it the next time it opens. Cancel
changes nothing. The dialog is reached from the ribbon button; the settings themselves are
reachable from a macro through the functions below, which is what a script should drive.

The group's fourth button, **Diagnostics**, sets nothing: it shows the modules
loaded into the process, the environment and the process counters, and is
described in the [README](../README.md#diagnostics-what-is-actually-loaded).

| Function | Does |
|---|---|
| `XRayXL_SetTraceParam(Source, Name, Value)` | Sets one setting |
| `XRayXL_GetTraceParam(Source, Name)` | Reads it back — omit either argument to get a table |
| `XRayXL_GetTraceSummary(Filter)` | Lists what has been traced, with live call counts |
| `XRayXL_IsArmed()` | `TRUE` while either source is armed |

`XRayXL_Arm` and `XRayXL_Disarm` start and stop a recording; `XRayXL_Disarm`
returns the number of rows dropped.

Everything is on by default except `BREAKPOINTS`, which changes the file's header. If you are after the most accurate timings, set
`ARGS`, `RETVAL` and `OBJECTS` to `FALSE` — each one costs work inside the call
being measured. The calling cell is always resolved and is not a setting: the VBA
tracer needs it to tell an error that escapes into a cell from one that propagates.

## Per-source settings

`Source` is `"XLL"` or `"VBA"`; omit it to address both at once. The two
sources are **independent** -- each holds its own value for every setting
below, and each is armed by its own value alone. Turning one off says
nothing about the other.

| `Name` | `Value` | Default | Meaning |
|---|---|---|---|
| `DEPTH` | `OFF` / `TOP` / `ALL` | `ALL` | Applies to **that source alone**: don't trace it / trace only the outermost call, the one Excel itself initiated / follow nested calls too |
| `ARGS` | `TRUE` / `FALSE` | `TRUE` | Capture argument values |
| `RETVAL` | `TRUE` / `FALSE` | `TRUE` | Capture return values |
| `OBJECTS` | `TRUE` / `FALSE` | `TRUE` | **VBA only** — `XLL`, or an omitted `Source`, is **refused**. Name an object argument or result, and describe a `Range`, `Worksheet` or `Workbook`. The one setting that calls Excel's object model from inside a traced call; `FALSE` renders every object as its address. An XLL argument is an `XLOPER` decoded structurally, with no object model to call, so there is nothing for it to gate |
| `BREAKPOINTS` | `TRUE` / `FALSE` | `FALSE` | **VBA only** — refused for `XLL` or an omitted `Source`. Adds one column, `breaks`, to the trace file: on each VBA exit row, how often that call paused in the VBA editor -- at a breakpoint, or at a `Stop` statement in the code -- so a duration that includes time in the debugger says so. Off by default because it changes the file's header ([TraceRowModel.md](./TraceRowModel.md)). In the Options dialog it is the check box on the Advanced page. Breakpoints and `Stop` are *handled* either way -- a call whose first lines have breakpoints still opens on time, and a `Stop` ends no frame -- this only adds the column |

**`DEPTH` changes which rows appear, never their shape.** Under `TOP` the
totals still count every frame, so the disarm report and
`XRayXL_GetTraceSummary` stay complete even when the file is deliberately thin.

## Settings for the recording as a whole

These take no `Source`.

| `Name` | `Value` | Default | Meaning |
|---|---|---|---|
| `BUFFERSIZE` | MB, e.g. `64` | `64` | Size of the ring buffer between the calculation thread and the writer; `0` writes every row synchronously |
| `BUFFERWHENFULL` | `PAUSE` / `DROP` | `PAUSE` | What a full ring does — make the calculation wait until it is half empty, or drop rows |
| `FORMAT` | `CSV` / `JSONL` | `CSV` | The trace file's format: CSV, with values as text, or JSON Lines, one object a line with every value structured and typed (see [the row model](TraceRowModel.md#json-lines)). Also **Options › Output › Format** |
| `LOGLEVEL` | `DEBUG`/`INFO`/`WARNING`/`ERROR` | `INFO` | The log's level (see [The log](#the-log)) |

The ring keeps file I/O off the calculation thread, so tracing disturbs the
timings as little as possible. The defaults suit most sessions.

- **`PAUSE`** keeps every row. When the ring fills, the traced threads wait
  until the writer has emptied it to half, so the calculation resumes with room
  for a burst rather than stalling again at once. The waits are counted as
  *hot-path pauses* on the disarm line.
- **`DROP`** never makes the calculation wait, and loses rows when the ring
  fills instead. It shows where through holes in the trace's `input` column;
  the count comes back from `XRayXL_Disarm` and appears on the disarm log line
  — the file itself carries no marker row. A dropped entry does not take its
  exit with it, so an exit row can appear with no entry.
- **`BUFFERSIZE=0`** writes each row the moment it happens. That is what you
  want when hunting the exact function that was running at a crash, because
  everything up to the crash is already on disk.

`BUFFERSIZE` accepts a unit — bare or `M`/`MB` is megabytes, `K`/`KB`
kilobytes. A ring below 16 KB or above 240 MB is refused. A row goes through the ring like any
other and waits or drops by `BUFFERWHENFULL`; one larger than the whole ring can never fit, so
it is dropped and counted even under `PAUSE`, since waiting for it would stall the calculation
for good. A value is at most 4 MB — an array past that is written as its shape alone — so with
the default ring only a deliberately small `BUFFERSIZE` meets this.

## Arming loads VBA if it is not already loaded

Excel loads VBA only when it needs it, and XRayXL patches the VBA interpreter
once, when you arm. So if you arm an Excel that has never opened a macro, there
would be nothing to patch and VBA would stay untraced for the rest of that
session -- even after you opened a macro workbook.

To avoid that, arming asks Excel to load VBA first, by reading the active
workbook's VBA project. Excel loads it exactly as it would when you open any
macro-enabled workbook. It happens only when you arm, and only when VBA tracing
is on: set `VBA` `DEPTH` to `OFF` and XRayXL will not touch VBA at all.

## When settings can change

**A refused call answers `#Err - ` and the reason**, and changes nothing. The
value is the return of `Application.Run`, so the caller reads it directly.

**Every setting is refused while armed** — disarm, set, re-arm — so a recording
can never change shape halfway through. `LOGLEVEL` is the one exception: it
affects only logging, never the trace, so it can be changed at any time.

On the ribbon this is why **Arm** is greyed while a recording is running and
**Disarm** while one is not, and why the capture settings in **Options** are
greyed while armed — the log level is not. It is the same rule, shown rather
than refused.

The setters also refuse when called from a cell. A trace setting changed
mid-calculation, or written by a formula, would no longer describe the run it
is attached to.

Both getters are volatile, so they can live in cells and keep up as you go:

```
=XRayXL_GetTraceParam()            ' every setting, both sources
=XRayXL_GetTraceSummary("*.xll")   ' Source | Module | Function | Calls
```

The summary's filter is a wildcard (`*` and `?`, any case) matched against the function
and the module, so `"*.xll"` picks out the add-ins and `"[Book1.xlsm]*"` one workbook.
When nothing matches it says `(nothing matches the filter)`; before anything is traced
it says `(nothing traced yet)`.

## The log

Alongside the trace, `%TEMP%\XRayXL\Logs\XRayXL_<pid>.log` records what the
*add-in* did — a different question from what your spreadsheet did. It is a
levelled log in the familiar Log4Net line format: timestamp, thread, level,
message.

```
2026-09-07 10:53:07,123 [ 4812] INFO  - armed 46 of 53 registered; declined 7 [module 0, proc 0, typetext 0, ours 7, detour 0, space 0]
2026-09-07 10:53:07,124 [ 4812] INFO  - arm cost: total 5ms = enumerate 0ms + name resolution 1ms + hook install 3ms
```

| Set it | How |
|---|---|
| At startup | `XRAYXL_LOGLEVEL=DEBUG` (or `WARNING`, `ERROR`) before launching Excel |
| Live | `Application.Run "XRayXL_SetTraceParam", "LOGLEVEL", "DEBUG"` — settable while armed |
| Developer detail | `XRAYXL_DIAG=1` before launching Excel, which adds diagnostics to the disarm report |

**Note the declines.** Every path where the tracer chose *not* to trace
something is counted and named — `module`, `proc`, `typetext`, `ours`,
`detour`, `space` in the line above. A tool that quietly skips things will
report a broken spreadsheet as a healthy one, so "we never looked" and "we
looked and found nothing" are never allowed to read as the same sentence.

## Environment switches

Read when Excel loads the add-in, so set them before launching Excel.

| Variable | Effect |
|---|---|
| `XRAYXL_LOGLEVEL` | Starting log level (see [The log](#the-log)) |
| `XRAYXL_OUTPUT_DIR` | Root for `Logs` and `TraceFiles` (see below) |
| `XRAYXL_CRASHDUMP=1` | Write a full minidump on a crash, as well as the text report |
| `XRAYXL_DIAG=1` | Diagnostics in the disarm report, and registers `XRayXL_FaultProbe`, which faults on purpose to test the crash handler |
| `XRAYXL_NOREGWATCH=1` | Do not watch for functions registered after arming; only functions registered at arm time are traced |
| `XRAYXL_RIBBON=0` | No ribbon buttons, and no COM object of ours in the process at all. Everything else is unaffected — it is the switch to reach for if you suspect the ribbon of anything |
| `XRAYXL_NOMESSAGEBOX=1` | No message box when the ribbon cannot load; the log still says why. A hidden Excel never gets the box |

While `%TEMP%\XRayXL\inert.on` exists, the add-in loads and does nothing at all —
no log, no commands, no hooks — which tells a crash caused by what it does from
one caused by its presence.

## Where the files go

| | |
|---|---|
| Trace | `%TEMP%\XRayXL\TraceFiles\XRayXL_Trace_<id>_<pid>.csv` — one per arm |
| Log | `%TEMP%\XRayXL\Logs\XRayXL_<pid>.log` |

XRayXL never deletes these. Each Excel process adds one log, each arm one trace
file, and each crash one report, so clear the folder yourself when you no longer
need them.

`XRAYXL_OUTPUT_DIR` moves all of it: set it before launching Excel and the
`Logs` and `TraceFiles` folders are created under that root instead of
`%TEMP%`. The test suites use this to keep each session's output separate.
