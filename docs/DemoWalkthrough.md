# XRayXL — Demo

A hands-on playground for XRayXL. Two ordinary demo add-ins and four macro-enabled
workbooks, so you can arm the tracer, run something, and read the timeline it produces.

> 64-bit Excel required. XRayXL traces 64-bit Excel only.

## What's in here

Everything you need is committed and ready to run — no compiler required.

```
dist/
  XRayXL64.xll             the tracer itself (the product)
  demo/
    DemoFinance64.xll    demo add-in: BlackScholes, PresentValue, CompoundReturn, SlowSum, Fibonacci
    DemoBehaviors64.xll  demo add-in: ReverseText, MakeSeries, MightDivide, SafeDivide, ThreadSafeSquare, CallCounter
    01_CalcChain.xlsm    a VBA UDF that calls nested VBA AND an XLL function
    02_Events.xlsm       Worksheet_Change, Workbook_Open, a button macro
    03_Errors.xlsm       an Err.Raise / On Error chain (threw / unwound / handled)
    04_Advanced.xlsm     a class module, recursion, and a UserForm + OnTime timer
```

The source those are built from lives here:

```
demo/
  src/                      the demo add-ins' C++ source (plain XLLs; they know nothing of XRayXL)
  src/*.vcxproj             the add-ins' projects, built by XRayXL.sln
  tools\Build-DemoWorkbooks.ps1   regenerates the workbooks (needs Excel)
  README.md                 this file; shipped into dist\demo\
```

The demo add-ins are deliberately boring, normal XLLs — they are the *add-ins you point XRayXL at*, not part of XRayXL.

## Load the three add-ins

**File ▸ Open** each of the three: `DemoFinance64.xll` and `DemoBehaviors64.xll`
from `dist\demo\`, and `XRayXL64.xll` from `dist\` one level up. Opening a `.xll`
loads it for the session (it runs the add-in's `xlAutoOpen`); you won't see a
window — it just registers its functions and commands. The two demo add-ins
provide the worksheet functions; `XRayXL64.xll` provides the tracer commands the
buttons call.

(Prefer them to load every time Excel starts? Add them instead via **File ▸ Options
▸ Add-ins ▸ Manage: Excel Add-ins ▸ Go ▸ Browse**. For a play session, File ▸ Open
is quicker.)

Expect your AV/EDR to notice `XRayXL64.xll` — it hooks Excel internals; that's what it does.

## Play with it

Every workbook has two buttons wired up for you:

1. Open a workbook (enable macros). The formula cells fill in from the demo add-ins.
2. Press **Arm XRayXL**.
3. Press the workbook's own button (**Recalculate**, **Run error chain**, **Run advanced**, …) — or just edit a cell.
4. Press **Disarm + trace**. A File Explorer window opens on `%TEMP%\XRayXL\TraceFiles`.
5. Open the **newest** `XRayXL_Trace_*.csv` — that's your timeline.

No VBA editor needed. (Under the hood the buttons call `Application.Run("XRayXL_Arm")` /
`("XRayXL_Disarm")`, which is also how you'd drive it from your own macros.)

The **XRayXL** buttons on the Developer tab do the same two things, and the
**Options** dialog holds the capture settings below — the workbook buttons
exist so the demo works whether or not the ribbon loaded, or the Developer tab is
showing. Arm from anywhere and all three agree: they are the same commands.

### What a trace row says

`kind` (entry/exit) · `source` (XLL/VBA) · `function` · `caller` (the calling cell) ·
`args` · `ret` / `rettype` · `depth` and `parent` · `ticks` (duration) · and
`outcome` (returned / threw / unwound / handled). One recalc of `01_CalcChain` shows,
in one file, a VBA UDF (`OptionBook`) calling the XLL `BlackScholes` and returning
1045.06 — VBA → XLL, nested, in a single activation.

## What each workbook is for

| Workbook | Press… | Watch for |
|---|---|---|
| **01_CalcChain** | Recalculate | `RiskWeighted`→`WeightFor` (nested VBA); `OptionBook`→`BlackScholes` (VBA→XLL); `SlowSum`/`Fibonacci` dominating the tick column |
| **02_Events** | edit the **Trigger** cell, or the button | `Worksheet_Change`, `Workbook_Open` — VBA that runs *outside* the calc engine; `CallCounter` (volatile) re-running each recalc |
| **03_Errors** | Run error chain | the `outcome` column: `E_Thrower`=threw, `E_Middle`=unwound, `E_Outer`=handled; the XLL `MightDivide` returning `#DIV/0!` next to `SafeDivide` |
| **04_Advanced** | Run advanced | a class (`Position`), recursion (`WalkTree`), a UserForm + `OnTime` timer; `MakeSeries` spilling an array; `ThreadSafeSquare` down a column running on several calc **threads** |

## The three output files

XRayXL writes only to `%TEMP%\XRayXL\`:

- **Trace** — `TraceFiles\XRayXL_Trace_<id>_<pid>.csv`, one per arm (created on the first traced row; a session that traces nothing leaves none).
- **Log** — `Logs\XRayXL_<pid>.log`, a levelled timeline (`INFO` by default) recording arming, the trace file's name, the process id and the loaded add-in path.
- **Crash report** — `Logs\XRayXL_crash_<pid>.txt`, a short text report (registers and module names, no workbook content) written if Excel crashes with the add-in loaded. A full minidump is off by default; set `XRAYXL_CRASHDUMP=1` before starting Excel to opt in (a dump of Excel contains workbook memory).

Nothing is written to your workbook — no cells changed, no VBA added, no source read.

### A few knobs (all via `Application.Run("XRayXL_SetTraceParam", …)`)

| Set | Effect |
|---|---|
| `"VBA","DEPTH","ALL"` / `"OFF"` / `"TOP"` | trace all VBA frames / none / only the outermost (the Arm button sets `ALL`) |
| `"XLL","DEPTH","ALL"` / `"OFF"` | same for the XLL side (on by default) |
| `"LOGLEVEL","DEBUG"` | more log detail (`DEBUG`/`INFO`/`WARNING`/`ERROR`); settable any time, even while armed |
| `"BUFFERSIZE", 64` | ring-buffer the trace, N MB, drained off-thread (0 = write synchronously) |
| `"VBA","OBJECTS",FALSE` | render objects as their address, with no calls into Excel's object model |

`Application.Run("XRayXL_GetTraceSummary")` lists what was traced with live call counts, and names the trace file.

For deeper start-up logging, set the environment variable `XRAYXL_LOGLEVEL=DEBUG` before launching Excel.

## Rebuilding

```powershell
msbuild XRayXL.sln /p:Configuration=Release /p:Platform=x64   # the add-ins -> build\x64\Release\
.\tools\Build-DemoWorkbooks.ps1                               # the workbooks -> dist\demo\
```

The add-ins are ordinary MSBuild projects, so they need nothing but the
compiler, and they build into `build\x64\Release\` with everything else.
**Regenerating the workbooks needs Excel**, with *Trust access to the VBA
project object model* enabled (Trust Center ▸ Macro Settings) — they are made
by driving Excel over COM and injecting VBA. That is why the finished `.xlsm`
files are committed: you only need this if you want to change the demo content
itself. Opening and using them needs nothing but ordinary macro-enabling.

`dist\` is assembled only by `tools\release.ps1`, which copies the built add-ins
in beside the workbooks. A normal build does not update it, so between releases
it holds the last released binaries — `dist\MANIFEST.txt` records the version,
tag and commit they came from.
