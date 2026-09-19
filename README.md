# XRayXL

**A flight recorder for the code running inside Excel.**

When a spreadsheet is slow, freezes, or crashes, XRayXL tells you which XLL
add-in and VBA functions actually ran — in what order, on which thread,
triggered by which cell or button, with what arguments, returning what, how
long each took, which functions threw errors and which handled them. It is a
window into internals Excel has never opened, and it takes the guesswork out
of diagnosing spreadsheet speed and stability problems.

XRayXL changes nothing about your workbook: no cells touched, no macros added,
no VBA source read. You arm it, recalculate or run your macros, then read the
trace file to see what really happened.

> **Warning.** XRayXL hooks into the internals of the Excel process —
> undocumented details Microsoft can change in any update. It does
> its best to adapt to different versions of Excel, but it is not fool-proof.
> Read [Before you run it](#before-you-run-it), especially if the machine is
> managed by someone else.

## How it works

Two tracers, one output.

- **XLL functions.** Excel's registration table names every registered add-in
  function and its address — a documented API — so XRayXL wraps each one with
  an inline detour that records the entry, calls the original, and records the
  exit. Functions registered after arming are picked up by watching Excel's
  registration callback. Nothing is derived; nothing is guessed.
- **VBA procedures.** The VBA interpreter reaches every opcode handler through
  one table of function pointers inside `VBE7.DLL`. XRayXL finds that table by
  its *shape* — not by a byte signature — and swaps a handful of entries for its
  own, so it sees each procedure start and stop without touching a byte of
  code. A per-thread shadow stack supplies depth and nesting; the calling cell
  comes from asking Excel.

**What it does not trace.** Excel's own built-in functions — `SUM`, `XLOOKUP`
and the rest — are deliberately out: XRayXL follows add-in and VBA code, not
the calculation engine's internals. COM and RTD add-ins are not traced either.
Everything else a calculation runs through is: XLL functions from any add-in,
VBA procedures, nested and recursive calls, and VBA outside the calculation
engine entirely — macros, buttons and event handlers — because the interpreter
hook is global.

## Example trace file
Here is a real trace of one recalculation, with VBA functions and an add-in's XLL
function called from four cells, in the order Excel calculated them:
`A3` calls `VbaOuter`, `A4` calls `TxE`, `A2` calls `VbaPlain`, `A1` calls `TxE`.
`VbaOuter` also calls the XLL function itself, through `Application.Run`.

```text
seq  input  kind   source  span  parent  depth  thread  qpc            module                         function  proc           typetext  caller  callerref                          argcount  args         ret  rettype  outcome   ticks  trust
1    1      entry  VBA     1     0       1      43652   3461532742964  [OneTimeline_39248.xlsm]Probe  VbaOuter  0x291330F7408  Double    cell    [OneTimeline_39248.xlsm]Sheet1!A3  1         a1:Double=3
2    2      entry  XLL     2     0       1      43652   3461532745498  TracedAddin64.xll              TxE       TxE            E         cell    [OneTimeline_39248.xlsm]Sheet1!A3  1         a1:E=3
3    3      exit   XLL     2     0       1      43652   3461532745611  TracedAddin64.xll              TxE       TxE                                                                                        3    Q        returned  113    exit
4    4      exit   VBA     1     0       1      43652   3461532745730  [OneTimeline_39248.xlsm]Probe  VbaOuter  0x291330F7408                                                                              4    Double   returned  2766   exit
5    5      entry  XLL     3     0       1      43652   3461532746066  TracedAddin64.xll              TxE       TxE            E         cell    [OneTimeline_39248.xlsm]Sheet1!A4  1         a1:E=4.5
6    6      exit   XLL     3     0       1      43652   3461532746182  TracedAddin64.xll              TxE       TxE                                                                                        4.5  Q        returned  116    exit
7    7      entry  VBA     4     0       1      43652   3461532746473  [OneTimeline_39248.xlsm]Probe  VbaPlain  0x290A38EFB18  Double    cell    [OneTimeline_39248.xlsm]Sheet1!A2  1         a1:Double=2
8    8      exit   VBA     4     0       1      43652   3461532746654  [OneTimeline_39248.xlsm]Probe  VbaPlain  0x290A38EFB18                                                                              6    Double   returned  181    exit
9    9      entry  XLL     5     0       1      43652   3461532746910  TracedAddin64.xll              TxE       TxE            E         cell    [OneTimeline_39248.xlsm]Sheet1!A1  1         a1:E=1.5
10   10     exit   XLL     5     0       1      43652   3461532747069  TracedAddin64.xll              TxE       TxE                                                                                        1.5  Q        returned  159    exit
```

<details>
<summary>Show the raw file</summary>

```csv
seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,caller,callerref,argcount,args,ret,rettype,outcome,ticks,trust
1,1,entry,VBA,1,0,1,43652,3461532742964,[OneTimeline_39248.xlsm]Probe,VbaOuter,0x291330F7408,Double,cell,[OneTimeline_39248.xlsm]Sheet1!A3,1,a1:Double=3,,,,,
2,2,entry,XLL,2,0,1,43652,3461532745498,TracedAddin64.xll,TxE,TxE,E,cell,[OneTimeline_39248.xlsm]Sheet1!A3,1,a1:E=3,,,,,
3,3,exit,XLL,2,0,1,43652,3461532745611,TracedAddin64.xll,TxE,TxE,,,,,,3,Q,returned,113,exit
4,4,exit,VBA,1,0,1,43652,3461532745730,[OneTimeline_39248.xlsm]Probe,VbaOuter,0x291330F7408,,,,,,4,Double,returned,2766,exit
5,5,entry,XLL,3,0,1,43652,3461532746066,TracedAddin64.xll,TxE,TxE,E,cell,[OneTimeline_39248.xlsm]Sheet1!A4,1,a1:E=4.5,,,,,
6,6,exit,XLL,3,0,1,43652,3461532746182,TracedAddin64.xll,TxE,TxE,,,,,,4.5,Q,returned,116,exit
7,7,entry,VBA,4,0,1,43652,3461532746473,[OneTimeline_39248.xlsm]Probe,VbaPlain,0x290A38EFB18,Double,cell,[OneTimeline_39248.xlsm]Sheet1!A2,1,a1:Double=2,,,,,
8,8,exit,VBA,4,0,1,43652,3461532746654,[OneTimeline_39248.xlsm]Probe,VbaPlain,0x290A38EFB18,,,,,,6,Double,returned,181,exit
9,9,entry,XLL,5,0,1,43652,3461532746910,TracedAddin64.xll,TxE,TxE,E,cell,[OneTimeline_39248.xlsm]Sheet1!A1,1,a1:E=1.5,,,,,
10,10,exit,XLL,5,0,1,43652,3461532747069,TracedAddin64.xll,TxE,TxE,,,,,,1.5,Q,returned,159,exit
```

</details>

Read row 1 as: *the VBA function `VbaOuter`, in module `Probe` of `OneTimeline_39248.xlsm`,
called from `[OneTimeline_39248.xlsm]Sheet1!A3`, declared `Double`, given
`3`.* The `caller` column says what **kind** of thing called it and
`callerref` describes which one — the same address Excel's own
`Range.Address(,,,True)` gives, so you can paste it straight back into a formula.

Rows 2 and 3 are the add-in call `VbaOuter` made: *`TxE`, exported by
`TracedAddin64.xll`, registered with argument code `E`, given `3`,
returned `3` in 113 ticks.* It still names `A3` as its caller, because
Excel reports the cell that started the chain. Each source counts its own
`depth`, so the add-in call reads `depth` 1 and `parent` 0; its place inside
`VbaOuter` shows in the order of the rows, between that function's entry and exit.
Row 4 closes `VbaOuter`: *returned `4`, a `Double`, after 2766 ticks.*

The other rows are the same shapes on their own: `VbaPlain` from `A2`
returning `6`, and `TxE` called directly from `A4` and `A1`.


### Errors say where they were thrown, and who caught them


```
VBA entry  E_Outer               depth=1 parent=0
VBA entry  E_Middle              depth=2 parent=1
VBA entry  E_Thrower             depth=3 parent=2
VBA exit   E_Thrower   threw     depth=3 parent=2  ticks=4657   trust=exit
VBA exit   E_Middle    unwound   depth=2 parent=1  ticks=6585   trust=exit
VBA exit   E_Outer     handled   depth=1 parent=0  ticks=10111  trust=exit
```
Every exit row carries an `outcome` (always `returned` for an XLL), and a VBA
chain reads outwards from the throw. `threw` is the procedure that raised. `unwound` is one the error passed
*through* — it ran nothing after the raise, so it did not handle it. `handled`
is the one that caught it. Everything else reads `returned`.

**Not supported:** asynchronous XLL functions (registered with a leading `>`
or an `X` handle argument) are marked in the trace but carry no duration — they
return before their answer exists, so there is nothing to measure. Real-time
data (`=RTD(...)`) is a COM mechanism rather than the XLL C API, and is not
traced at all.

Details of the Trace File in [docs/TraceRowModel.md](./docs/TraceRowModel.md).


## Getting started

**No build required.** `dist/` is committed, so the
[latest release](../../releases/latest) — or a plain clone — already contains a
working tool:

```
dist/
  XRayXL64.xll               the add-in
  LICENSE                  GPL-3.0
  THIRD-PARTY-NOTICES.txt
  MANIFEST.txt             version, tag, commit and SHA256 of every file here
  demo/                    two demo add-ins + nine macro-enabled workbooks
```

Point Excel at `dist\XRayXL64.xll`: either add it permanently through
File → Options → Add-ins → Manage: Excel Add-ins, or drag the `.xll` onto an
open Excel window to load it for that session only.

An **XRayXL** group appears at the far right of the **Developer** tab, with three
buttons: **Arm**, **Disarm** and **Options** — the last opens a dialog holding the
capture settings. Two things to know: Excel hides the Developer tab by default
(File → Options → Customize Ribbon, tick *Developer*), and the buttons appear once
a workbook is open, not on Excel's start screen. Each is also a
registered command, so nothing needs the ribbon — see
[Trace something](#trace-something). If the buttons cannot be loaded, XRayXL says so
and carries on working; [Before you run it](#before-you-run-it) explains when
that happens.

[**`dist/demo/`**](docs/DemoWalkthrough.md) is a guided tour in nine workbooks: a first
trace, argument values, callers, errors, a real VBA yield-curve model, classes and
objects, the call tree, threads, and an option book built from XLL functions. Load the
two demo add-ins with File → Open, open a workbook, press **Arm** on the ribbon, do
what the sheet says, press **Disarm**, and read the trace.
[`docs/DemoWalkthrough.md`](docs/DemoWalkthrough.md) is the walkthrough.

`dist/` is written only by `tools\release.ps1`, and only after a full test
sweep passes, so it holds what a release shipped rather than whatever was last
compiled. Between releases it lags the source beside it; `MANIFEST.txt` says by
how much.

The trace lands in: `%TEMP%\XRayXL\TraceFiles\XRayXL_Trace_<id>_<pid>.csv`
The log lands in: `%TEMP%\XRayXL\Logs\XRayXL_<pid>.log`

## Build

If you want to build it yourself you'll need 64-bit Windows and **Visual Studio Build Tools**
with the *Desktop development with C++* workload, which brings the MSVC
toolset (v145), MASM and the Windows SDK.

```powershell
msbuild XRayXL.sln /p:Configuration=Release /p:Platform=x64
tools\deploy.ps1        # copies the XLL to build\addin\
```

That builds into `build\x64\Release\`: the add-in itself, the `TracedAddin` the
suites need, the two demo add-ins, and the unit tests under `tests\sweep\unit\`. A normal build never
writes `dist\` — only `tools\release.ps1` does that, after a green sweep.

The output is one native DLL — `build\x64\Release\XRayXL\XRayXL64.xll` — with no
runtime dependencies beyond Windows itself. No .NET, no installer, and nothing
to register permanently: Excel serves ribbon controls only to a COM add-in, so the
XLL is one for as long as the connect takes, then deletes its own registration
(see [Before you run it](#before-you-run-it)).


### Trace something

Press **Arm** in the XRayXL group on the Developer tab, recalculate or run your
macros, then press **Disarm**. Or, from VBA or any automation client, with no window and no focus:

```vba
Application.Run "XRayXL_Arm"          ' start recording

Application.Calculate                 ' ...or press F9

Application.Run "XRayXL_Disarm"       ' stop, flush, close
```

The two are the same thing: the buttons call these commands. Arm from a macro
and the ribbon follows it; change a setting in **Options** and
`XRayXL_GetTraceParam` reports it. While armed the capture settings are greyed —
they are read once, at arm, so they are refused until you disarm.


### Choosing what to capture

Registered functions, called through `Application.Run`, decide what is recorded
and report on the session: `XRayXL_SetTraceParam`, `XRayXL_GetTraceParam`,
`XRayXL_GetTraceSummary` and `XRayXL_IsArmed`. You can turn each source off,
limit a source to the top-level call, drop argument, return-value or object
capture for tighter timings, and size the ring buffer that keeps file I/O
off the calculation thread.

Everything is on by default and the defaults suit most sessions.
**[docs/TraceOptions.md](./docs/TraceOptions.md)** is the reference — every
setting, what it costs, when it can be changed, and how to read the log.

## Before you run it

Honest caveats, in roughly the order they will matter to you.

- **This hooks Excel's internals.** It patches function entry points in a live
  Excel process. Every hook is fault-guarded, the VBA hooks stand down after
  repeated faults, and the test suite exists to keep it that way — but the blast radius of a bug is
  somebody's Excel, so treat it accordingly.
- **Antivirus and EDR will plausibly notice.** Inline hooking of `excel.exe` is
  malware-shaped behaviour. On a managed or corporate machine, expect a
  conversation with whoever owns security before it runs at all.
- **Check it against your licence terms.** Instrumenting Office internals may
  sit in tension with the Office licence. See [Legal](#legal), and if in doubt
  ask legal or compliance before running it.
- **The trace file contains your data.** Argument values, return values, cell
  references and function names all land in it — so whatever your workbook
  computes is in there, in the clear. It is a plain file in `%TEMP%`; treat it
  as sensitive as the spreadsheet itself.
- **The ribbon buttons make XRayXL briefly a COM add-in.** Ribbon controls cannot
  be served any other way. To load it, XRayXL writes a CLSID, a ProgId and an Excel
  add-ins entry under `HKEY_CURRENT_USER`, connects itself, and **deletes all
  three immediately** — they are not left behind for the session, and the
  add-ins entry is written `LoadBehavior=0`, so nothing of it can autoload into
  your next Excel. If any of that is refused — a locked-down profile, an
  elevated Excel, or Excel itself declining — you get a message box saying so,
  no buttons, and an add-in that otherwise works normally. `XRAYXL_RIBBON=0` skips
  the whole thing, and is the switch to reach for if you suspect the ribbon of
  anything.
- **Crash dumps are off by default.** A crash always writes a small text report
  (registers and module names, no workbook content). A full minidump — which
  *would* contain workbook memory, and whose writer is an EDR signature — is
  opt-in via `XRAYXL_CRASHDUMP=1`, and every session logs which mode it is in.
- **64-bit Excel only.** Not for a deep reason — the two hooking thunks are
  hand-written x64 assembly, and the argument and return-value decoders read
  the x64 calling convention directly. 32-bit is possible in principle, but it
  is a port and a re-measurement, not a build flag.
- **Durability across Excel updates is unproven.** Nothing is hardcoded: every
  address is derived at runtime, which is *designed* to survive an Excel
  update — but only the VBA dispatch-slot indices have been checked across
  many builds (41 VBE7 builds, 2012 to 2026); the tracer as a whole is tested
  only against the Excel its tests run on. If an
  update breaks a derivation, the add-in refuses to arm and says so, rather
  than producing a wrong trace.

## Running the tests

The tests drive real Excel through **StretchXL**, a general-purpose Excel test
manager written for this project and kept independently useful. They need
64-bit Excel installed and an interactive desktop session.

```powershell
msbuild XRayXL.sln /p:Configuration=Release /p:Platform=x64
.\StretchXL\StretchXL.ps1 -Parallel 8 -Path .\tests\sweep -OutDir <any folder>
```

Every asserted formula lives in a workbook the test **saved, closed and
reopened** first, and the cases are grouped by *how the calculation was
triggered* — F9, an edit, a button, opening the file. Neither is fussiness:
a formula loaded from a file and one assigned in memory are different things
inside Excel's calculation engine, and so are the ways a recalculation starts.
[docs/Testing.md](./docs/Testing.md) has the prerequisites, how to read a run,
and the principles behind that; [tests/sweep/README.md](./tests/sweep/README.md) is the
inventory of what each suite defends.

## Where to go next

| Document | What it covers |
|---|---|
| [docs/DemoWalkthrough.md](./docs/DemoWalkthrough.md) | A guided tour — demo add-ins and nine workbooks, each showing one part of the trace |
| [docs/TraceOptions.md](./docs/TraceOptions.md) | Every capture setting, what it costs, and the add-in's log |
| [docs/Implementation.md](./docs/Implementation.md) | How it works — the vehicle, what is derived, and the two tracers |
| [docs/VBATracing.md](./docs/VBATracing.md) | A layered walkthrough of VBA tracing — what is patched, the shadow stack, return values and error outcomes, for a general programmer |
| [docs/TraceRowModel.md](./docs/TraceRowModel.md) | The trace file's contract, in full |
| [docs/Testing.md](./docs/Testing.md) | How the tests are run, and the principles that decide whether a suite is worth running |
| [tests/sweep/README.md](./tests/sweep/README.md) | What each suite defends, and how to add a test |


## Acknowledgements

The VBA-tracing approach is inspired by
**[Dr. D. Azzopardi's `desva/VBA` project](https://github.com/desva/VBA)** and a
conversation in a Canary Wharf bar on a warm London evening in 2016 — a working
proof that Excel's VBA execution can be traced by locating the p-code
interpreter's opcode dispatch table and patching it to emit timed, attributed
records, without reading any VBA source. That project demonstrated the single
hardest mechanism this one depends on. **XRayXL adopts the *approach* and
builds an independent implementation; it reuses none of that code.**

## Legal

XRayXL is an independent project. It is not affiliated with, endorsed by or
supported by Microsoft. Microsoft, Excel, Office and VBA are trademarks of
Microsoft Corporation, used here only to say what XRayXL works with.

Its purpose is interoperability: letting you observe the add-in and VBA code
running in a copy of Excel you are licensed to use. It works in memory, inside
the Excel process it is loaded into. It does not modify Office files on disk,
bypass licensing or activation, or read VBA source, and it redistributes no
Microsoft code other than the Excel SDK header listed under
[Third-party code](#third-party-code).

You are responsible for making sure your use complies with your Office licence
and the law where you are, and for using XRayXL only on workbooks and add-ins
you are entitled to inspect. If in doubt, ask your legal or compliance team
before running it.

## Licence

Copyright © 2026 Andrew Lockhart.

XRayXL is free software: you may redistribute it and modify it under the terms
of the **GNU General Public License, version 3**, as published by the Free
Software Foundation. It is distributed in the hope that it will be useful, but
**without any warranty** — without even the implied warranty of merchantability
or fitness for a particular purpose. See the [LICENSE](./LICENSE) file for the
full terms.

Contributions are accepted under the same licence.

You may use, modify and run XRayXL freely, including inside a company. If you
distribute a modified version, or anything built on it, that work must be
released under the GPL too — so it cannot be taken closed-source. The full
terms are in the licence; that sentence is not a substitute for them.

### Third-party code

XRayXL includes two components under their own terms, both reproduced in the
files named:

| Component | Licence |
|---|---|
| **MinHook** — the inline hooking library, © 2009-2017 Tsuda Kageyu | BSD 2-Clause, `src/third_party/minhook/LICENSE.txt` |
| **`xlcall.h`** — the Excel C API header, from Microsoft's Excel XLL SDK | Microsoft's Excel 2013 XLL SDK terms; provenance in `src/third_party/xlcall.h` |

MinHook's licence requires its copyright notice to travel with **binary**
redistributions as well as source, so any release of `XRayXL64.xll` carries it
alongside.
