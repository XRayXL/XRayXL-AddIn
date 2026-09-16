# Testing

The front door: what a test run needs, how to run one, how to read the result,
and the principles that decide whether a suite is worth having. Nothing here is
repeated from the documents it points at.

| | |
|---|---|
| [`tests/sweep/README.md`](../tests/sweep/README.md) | **The suites.** Layout and conventions, the inventory of what each suite defends, and the rules every driver keeps. Read before writing a test. |
| [`StretchXL/StretchXL.md`](../StretchXL/StretchXL.md) | **The manager.** Excel lifecycles, session modes, deadlines, the dialog watchdog, measured close, dumps and exit codes. Read before touching the harness. |
| [`TraceRowModel.md`](./TraceRowModel.md) | **The contract the tests assert against**, enforced by the suites' single reader, which throws on any violation. |

---

## What you need

- **64-bit Excel, installed.** The tests drive the real thing. There is no mock.
- **An interactive desktop session.** Excel is started visibly and the harness
  dismisses dialogs; it will not work over a service account or a headless
  build agent.
- **The product, built and deployed** — `tools\deploy.ps1` puts it where the
  suites expect it.
- **`TracedAddin64.xll`** — an ordinary add-in that exists to be traced, standing
  in for the third-party XLLs a user would have. The suites that register it —
  `breakit`, `integration`, `modes`, `timeline`, `xll` — will not run
  without it. Two of its functions exist for the thread-safe soak: `TxRowCol`
  takes doubles, and `TxRowColQ` takes and returns `XLOPER12`, allocating each
  result for Excel to free through `xlAutoFree12`. `msbuild XRayXL.sln` builds
  it along with everything else, or on its own:

```powershell
msbuild tests\fixtures\TracedAddin\TracedAddin.vcxproj /p:Configuration=Release /p:Platform=x64
```

- **The unit tests, built.** `msbuild XRayXL.sln` builds them too, into
  `build\x64\Release\unit\`; the `unit` suite will not run an exe older than
  the source it came from.

The harness starts and closes Excel processes of its own throughout a run. It
kills only the ones it created, by process id — but do not run a sweep on a
machine where you are also working in Excel.

## Running it

```powershell
# everything
.\StretchXL\StretchXL.ps1 -Parallel 8 -Path .\tests\sweep -OutDir <any folder>

# the control: no -Path, so no tests -- proves the harness clean on its own
.\StretchXL\StretchXL.ps1 -Parallel 8 -OutDir <any folder>

# one suite, one trigger family, one test -- the folder is the selection
.\StretchXL\StretchXL.ps1 -Parallel 8 -Path .\tests\sweep\stress\change -OutDir <any folder>

# the order-dependency soak: dirty reused sessions, shuffled, replayable by seed
.\StretchXL\StretchXL.ps1 -Parallel 8 -Runs 20 -SessionMode Reuse -RandomOrder -Path .\tests\sweep -OutDir <any folder>

# (tests that load the same add-ins are kept together, so each session lives through
#  its group; -NoGroupBySession mixes every test into every session instead)

# the close-down regression: the interactive (X click) arm, warmed up first
.\StretchXL\StretchXL.ps1 -Parallel 8 -Path .\tests\sweep -OutDir <any folder> -CloseWithX -Warmup
```

`-Parallel` is mandatory; `-OutDir` is created if it does not exist. The
**floor run** — no `-Path` — starts Excel, adds a workbook and closes it. Run it
when a failure looks like it might be the harness rather than the product.

## Reading a run

One line per result, to stdout, as it completes:

```
worker=1 test=proof/binding.test.ps1 run=1 pid=34612 verdict=PASS outcome=CLEAN shutdown=2.3s test_s=1.1s elapsed=4.5s t=4.5s cases=3/3
```

**Two axes, always both reported.** `verdict=` is what the test concluded —
`PASS`, `FAIL`, `SKIP`, `TIMEOUT`, `ERROR`. `outcome=` is what the Excel session
did — `CLEAN`, `CRASH`, `HANG`, `EXITED`, `FAILED`. They are independent:
`verdict=PASS outcome=HANG` is a real line this harness produces, and it means
the test was right and Excel still would not close. That is a result, not a
contradiction.

**The exit code is the gate.** StretchXL exits `0` only if every result was
benign. Any failing verdict, any bad outcome, or fewer results than expected
exits `1` — so CI can gate on it without parsing anything.

Where things land:

| | |
|---|---|
| `<OutDir>\results-<timestamp>.jsonl` | One JSON line per result, with the test's captured output. The **first line records the whole configuration**, so a result set explains how it was produced |
| `<OutDir>\work\session_w<worker>_<n>\` | Each session's add-in output — its log, trace files and diagnostics — redirected there by `XRAYXL_OUTPUT_DIR`, never under a pid that `%TEMP%` will reuse |
| `%TEMP%\StretchXL\` | The harness's own ledgers and scratch |

A run ends with counts per outcome, crash and hang rates **with their
denominator**, and min/p50/p90/max of the measured shutdown times.

## Three principles

**1. Triggers and formula state are not optional dimensions.** Excel's ways of
starting a calculation are not equivalent inside the engine, and neither are the
ways a formula got into a cell. A suite that only calls `Application.Calculate`
on formulas it just assigned **will report a broken tool as healthy** — that
happened here, for months. So every asserted formula lives in a workbook that
was saved and reopened (**`AsLoaded`**), and the triggers are covered as
first-class cases: F9, `CalculateFull`, edit-and-enter, open, buttons, events,
ActiveX, selection. `tests/sweep/stress/` is grouped by trigger kind for exactly this
reason.

**2. A degraded rung is normal for the product and fatal for a test run.** The
tracer is built to degrade rather than guess — arm with types off, decline a
column it cannot establish, write empty rather than wrong. That is correct in a
user's session and worthless in a suite, where a test that silently ran
against a half-armed tracer is a green light for a broken build. The harness
refuses to hand back a session unless arming actually succeeded.

**3. Test the export shapes, not the frameworks.** An XLL's export address is
either the function's own code or a six-byte `jmp qword ptr [rip+disp32]` into a
writable data slot, and the tracer must handle both. Nothing about that is
specific to any framework — it is a property of the binary: Excel-DNA produces
the indirect shape, and so do incremental linking and import thunks.
`TracedAddin` therefore exhibits both itself. `jumptable.asm` supplies genuinely
indirect exports, and incremental linking is off so the rest are genuinely
direct. It also carries the case where the registered name differs from the
export name — Excel-DNA's `f0` shape — which is the only case that can tell a
name looked up from a name merely echoed.

**The gap this leaves is managed code.** No Excel-DNA add-in is built, committed
or loaded anywhere in this repository, and no suite would load one if it were.
The binary shapes are covered; a real CLR-hosted add-in — its load-time
registration, its marshalling layer, its own threading — has never been traced
by these tests. That is a known gap rather than an oversight, and it is the one
to close before claiming Excel-DNA support specifically.

Under those, the type codes that actually bite are worth naming, because they
are where the XLL decoder has been wrong before: a `%` string, a `K%` array, an
`O`, a `P`/`R` return, an async `X`.

## Adding a test

**One self-contained file per case.** A test carries its own case data — a
`$case = @{...}` literal and a call into its suite's `_driver.ps1` — so the
directory listing *is* the inventory: there is nothing to regenerate and no way
for a case to silently not exist. Files whose names start with `_` are never
discovered as tests. [`tests/sweep/README.md`](../tests/sweep/README.md) has the layout,
the per-suite configuration and what each suite is defending.

**A test binds to the session it is given, by window handle, and does exactly
one thing.** It never starts, closes, kills or times an Excel — the manager owns
every lifecycle, and that is what makes `-Parallel` safe and a hang
distinguishable from a crash.

**A unit test is a folder under `tests/sweep/unit/`.** It holds a C++ program
that prints `[PASS]`/`[FAIL]` lines and exits non-zero on a failure, a
`.vcxproj` that imports `..\UnitTest.props` and compiles the product sources it
tests, and a two-line `.test.ps1` that calls `Invoke-UnitTest`. Add the project
to the `unit` folder of `XRayXL.sln`, and its exe to `RequireNotOlderThan` in
`unit/suite.psd1`.

**`tests/` is organised by how each thing runs.**

| Folder | What it holds |
|---|---|
| `tests/sweep/` | Every test, including `unit/`, which runs the product's own code with no Excel. A sweep runs all of it, and `tools\release.ps1` will not publish unless it is green. |
| `tests/instruments/` | Long-running measurements, run one at a time and never in a sweep: `vbahammer/`, a crash hunt that may kill its Excel, and `tracesoak/`, a leak and cost measurement, with `_instrument.ps1` holding what they share. |
| `tests/fixtures/` | What the tests load but are not tests: `TracedAddin/`, the add-in above, and `xll_common/`, the header it shares with the demo add-ins. |

The instruments sit outside `tests/sweep/` because StretchXL runs every
`*.test.ps1` under the folder it is given; the folder a file is in is what decides
whether a sweep runs it.
