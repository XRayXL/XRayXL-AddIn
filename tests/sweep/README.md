# The sweep: XRayXL's test suites (run by StretchXL)

Every `*.test.ps1` under this folder is one test, and a sweep runs them all —
`tools\release.ps1` will not publish unless it is green. This file is the inventory:
what each suite defends, how a test is laid out, what a machine must have
before any of it can pass, and the rules every driver keeps. **How to run them
is in [docs/Testing.md](../../docs/Testing.md)** — prerequisites, the commands,
and how to read a result.

## Layout and conventions

```
tests\sweep\
  suite.psd1          root config: ModulesOfInterest for dump triage,
                      RequireNotOlderThan (refuses to start against a stale
                      deployed add-in), SessionEnvironment (every session
                      writes into its own directory)
  _xray_common.ps1    shared product plumbing: log and trace paths by pid,
                      XRayXL_Arm / XRayXL_Disarm via Application.Run, the
                      trace-mode setters, totals and summary parsing,
                      New-XRayMacroBook (build, save, reopen; one book or
                      several), Invoke-XRayFormulaTrace and
                      Invoke-XRayArmedSession (the arm, recalculate, disarm
                      cycles), the polling waits, locale-free cell values,
                      the row accessors every test reads a procedure's row
                      with, and Read-TraceFile -- the ONE reader of the trace
                      file, enforcing docs/TraceRowModel.md: it THROWS on any
                      contract violation
  <suite>\
    suite.psd1        XLLs to register per session, settle time, timeout override
    _driver.ps1       the suite's driving logic (one function)
    ...\*.test.ps1    the listing IS the inventory -- each file is its whole test
```

Underscore-prefixed files are never discovered as tests. **Most test files
carry their own case data** — a `$case = @{...}` literal followed by a
call into the suite's `_driver.ps1` — so editing a test is editing its file,
and there is nothing to regenerate and no way for a case to silently not
exist. The rest are self-contained scripts that drive Excel themselves,
report several named cases with `Check`, and finish with one verdict; they are
the tests whose subject is a sequence (arm, re-arm, set a parameter, read it
back) rather than one call.

The `$StretchCollectOnly` guard near the top of a case file lets the aggregate
tests (`xll\all-cases-one-sheet`, `vba\co-compiled`) harvest every `$case` by
dot-sourcing the files. Each test gets an Excel with the suite's XLLs already
registered — its own fresh one by default, or a shared long-lived one under
`-SessionMode Reuse`/`ReuseClean` (the process, its settings and XLL state
persist for realism; workbooks are always closed between tests, and ReuseClean
also resets settings); arming is the test's own explicit act either way.

## The suites

| Suite | What it defends |
|---|---|
| `xll\` | What the trace SAYS about one XLL call, against answers derivable by hand. Numbered files under `cases\`, one formula each, asserting the cell value, the decoded `args` and the decoded `ret` (several cases share a function on purpose, so the numbered FILE is the case's identity); plus `all-cases-one-sheet` (every case in one armed burst), `caller-is-an-array-range` (a CSE range as the caller, and an array RETURN from one cell — two different things), `derivation` (arming reports what it armed, and the VBA dispatch-table derivation agrees with whether VBE7 is loaded), `late-registration` (an XLL registered AFTER arming is still hooked), `register-then-use` (registered and called with nothing in between), and two soaks, `threadsafe-soak-double` and `threadsafe-soak-xloper` (a thread-safe function in 20,000 cells recalculated 20 times on 8 threads, every cell value and every trace row checked against the cell's own address; the XLOPER12 one also answers out-of-range arguments with a message naming the bad one, and returns a result Excel frees) |
| `vba\` | Identity, the shadow stack, nesting, recursion, error unwinds and frame layouts. Files under `cases\` state a VBA workload and the arithmetic the totals must satisfy (statements, procedures, depth, frames opened == closed, per-frame `outcome`); bespoke tests take the ones that need a class module, a form, two workbooks, a spilled array or a second arming session, including `co-compiled` (every case's setup in one module) and `return-value-from-cell` (every Function's result at every depth — scalars, Variants, five array shapes, Subs and Property Lets that must report NO value — against the values Excel computed) |
| `stress\` | A saved workbook per case — events, buttons, ActiveX, class modules, sheet code, cross-workbook and .xlam dependencies — grouped by trigger kind: `run\`, `calc\`, `change\`, `button\`, `activex\`, `open\`, `select\`. Every case additionally has to satisfy the invariants that hold whatever it does: frames opened == frames closed, no hook faults, the circuit breaker still closed, no procedure lost to a full table |
| `breakit\` | Deliberate attempts to break the decoder: hostile strings, surrogate pairs, CSV-hostile quoting and commas, extreme numerics, integer minima, re-entry and genuine nesting (the files under `fuzz\`), plus `run-evaluate` (Run and Evaluate calls with no calling cell). The unambiguous wrongs always fail; the cases whose right answer is derivable by hand also assert the decoded value, and the rest report it |
| `modes\` | The trace-mode surface: `XLL DEPTH=OFF` hooks nothing, `VBA DEPTH=TOP` emits only the depth-1 frame while the totals still count every frame, `VBA DEPTH=OFF` leaves the dispatch table untouched, ARGS and RETVAL change what is emitted and never what is counted, setters refuse while armed and from a cell, LOGLEVEL gates the log, `XRayXL_IsArmed` tracks both sources, the summary counts per arming session, a re-arm leaves no stale frames, and every setting reads back in all four query shapes |
| `ring\` | The buffered output path: an ample ring loses nothing and keeps file order, `PAUSE` throttles the producer rather than dropping, and a starved ring drops but ACCOUNTS for every drop — rows written + rows dropped == frames opened + closed, with the loss reported by `XRayXL_Disarm` and locatable as holes in the `input` column, never as a row in the CSV |
| `timeline\` | One file, one monotonic sequence, both sources, every row named and attributed to its calling cell, XLL nesting inside VBA — single-threaded, and again under forced multithreaded calculation where write order and event order genuinely diverge |
| `integration\` | A UserForm with a timer, end to end: VBA tracing is not confined to UDFs reached from a recalc, and form event code itself appears |
| `format\` | The trace-file contract itself ([docs/TraceRowModel.md](../../docs/TraceRowModel.md)): the shared reader must REFUSE a renamed or reordered column, an unknown kind, a seq inversion, a reused `input` — refusal paths the product tests never exercise |
| `unit\` | The product's own code, with no Excel: one folder per test holding a C++ program built by `XRayXL.sln` from the shipping sources, and the `.test.ps1` that runs it and reports each check as a case — the caller decoder, the registration-string parser, the trace-parameter grammar, the value decoder, the procedure table, the p-code scan guard and slot roles, the output ring under contention, the row formatter, and closing a session under writers |

Every `*.test.ps1` is one test, so the directory listing is the inventory:
`find tests/sweep -name "*.test.ps1"` lists it, and nothing here counts it.

## What a machine needs before any of this can pass

These are preconditions, not options. A machine that does not meet one will
show a suite red for a reason that is not a defect, so check here first.

- **Any UI locale.** Cells are read through `Range.Value2` and rendered by
  `Get-XRayCellText` the same way everywhere — a `.` decimal separator,
  `TRUE`/`FALSE`, errors named from their CVErr code — and VBA decimals are
  built by arithmetic rather than parsed from a string, so a comma-decimal or
  translated Excel gets the same answers.
- **64-bit Office.** `LongLong` and the `^` type suffix are 64-bit-only VBA and
  appear across the parameter-type cases; a 32-bit host will not compile them.
- **An interactive desktop session.** Buttons, ActiveX controls, UserForms and
  the modal-dialog cases need a real window station. Under a service or a
  locked-down session the MSForms designer refuses to insert a control, and
  those tests SKIP rather than fail (see below).
- **Dynamic-array spill.** `vba\dynamic-array-spill-udf` and the array-returning
  stress cases assume a build where an array result spills into neighbouring
  cells rather than being trimmed to one cell.
- **VBA project access trusted** (File > Options > Trust Center > Macro
  Settings > Trust access to the VBA project object model). Almost every VBA
  test builds its module through the VBIDE object model; without it they SKIP.
- **A deployed add-in that is not older than the build it came from.** The root
  `suite.psd1` refuses to start otherwise — a green sweep against a stale
  binary is worse than no sweep.
- **Unit tests built from the current source.** `msbuild XRayXL.sln` builds them
  into `build\x64\Release\unit\`, and `unit\suite.psd1` refuses to start against
  an exe older than the product source or its own test source.
- **One structural constant is pinned:** `xll\derivation` asserts the VBA
  dispatch table derives to `slots=1700`, the table's size. How many distinct
  handlers fill it differs between VBE7 builds, so only a positive `distinct`
  count is required.

`SKIP` means the machine cannot run the case at all: VBA project access is not
trusted, this Excel will not insert an ActiveX control, this Excel will not add
a UserForm component. Nothing about the PRODUCT skips. A raise slot that fails
to verify, a setter that does not take effect, a trace with a dropped row in
it — all of those fail. The one other SKIP is `xll\register-then-use` when
Excel binds the name to the original registration: the call that file exists
to watch never happened, though everything else it checks must still hold.

A line beginning `observed` in a test's output records a measurement with no
right answer to hold it to, such as an undocumented Excel behaviour. It is
never counted as a case, so it cannot pass or fail.

## Rules every driver keeps

- **`AsLoaded` always.** Every asserted formula lives in a workbook that was
  saved and reopened — a formula assigned and calculated in the same breath
  does not exercise the same engine path. `New-XRayMacroBook` in
  `_xray_common.ps1` is the one place that sequence lives.
- **Refuse a lossy trace.** Disarming returns the number of rows the run
  dropped; `Stop-XRayTrace` reads it and fails the test when it is not zero, so
  a starved buffer says "trace incomplete" instead of surfacing three lines
  later as a missing frame.
- **Poll, never guess a wait.** `Invoke-XRayRecalc` and `Wait-XRayCalcDone`
  wait on `CalculationState`, disarming returns once the trace file is closed,
  and anything else waits for its own condition through `Wait-XRayCondition`.
- **Reset at the start, never restore at the end.** `Set-XRaySessionDefaults`
  disarms a session an earlier test left armed -- and fails the test if it
  cannot, since every setting is refused while armed -- then puts every trace
  mode, the buffer, the log level, events and calculation threading back to the
  default before each test, so a test sets what it needs and leaves it. An
  earlier test's trace files are moved to `TraceFiles\earlier\`, not deleted.
- **A file a test leaves loaded gets a per-run name.** An add-in copy stays
  loaded after the test, so under `-SessionMode Reuse` a repeat of the same test
  must not reuse its name.
- **A test whose path did not run SKIPs, it does not PASS.** Under `-SessionMode
  Reuse` an earlier test can do a test's set-up for it: a function already
  registered is hooked at arm, not late, and a later arm releases the table a
  first arm measures. A test that depends on a first event checks for it and
  skips with the reason, and a test of a path asserts the evidence the path ran
  (a `late arm: hooked` line, a rise in the declined count), not only an outcome
  that another path could also produce.
- **Bind by handle, never by guess.** No `SendKeys`, no `GetActiveObject`;
  sessions are addressed by the hwnd the manager designates.
- **AutoRecover off** on every workbook object, including reopened ones.
- **Never modify a user file.** Everything is built under the session's work
  directory.
- **Kill nothing.** The manager owns every lifecycle.
- **Close your own leftovers, by leaf.** A test that saves and reopens a
  workbook closes its own earlier copy first (`Close-OwnLeftover`): under
  `-SessionMode Reuse` the pid-keyed name is the same for every test in the
  session, and a still-open predecessor makes `SaveAs` throw — 27 failures in
  one shuffled sweep, every one the second test of a family in its session.
  The rest of the session's dirt stays; that is the realism the mode exists for.
- **Expected rows and values come from what the test did**, never from the
  tracer's own counters: two counts of the same thing can be wrong together.
  A case says so in one of two ways, and both are checked against the trace:
  - **`Calls`** -- every call of that source, in entry order, as
    `@{ Function; Args; Ret; RetType; Outcome; Depth; Caller; Cell; Parent }`.
    `Parent` is the index in the list of the call this one ran inside, or -1 for
    none; an expected value ending in `{` is a prefix, for an array whose shape
    is the point. `Test-ExpectedTrace` requires exactly those rows: an extra or a
    missing one fails on the count before anything else is compared.
    **`CallsAnyOrder`** when the order is Excel's -- a sheet of cells -- so each
    expected call must match a different row, and `Parent` is not used.
  - **`Counters`** -- the name of a macro in the case's own VBA returning
    `"Name=N;Name=N"`, for the cases where how often something runs is Excel's to
    decide (a recalculation, an event cascade, a volatile UDF). The driver reads
    it BEFORE arming and after disarm and uses the difference, so the tally covers
    the same window as the trace and adds no rows to it;
    `Test-TracedCallsMatchCounters` then requires a row per counted call.
  - **`DepthCappedUnder`** -- the index of the deepest recorded call, for a
    recursion past the shadow stack: exactly one `depth-capped` row must name it.
- **Assert the row invariants on every row** (`Test-RowInvariants`): every entry
  has its exit, one of each per span, ticks equal to the exit's qpc less the
  entry's, a same-source parent one level up, and a function named. Also assert
  the expected caller of the first traced frame, keyed on how the test
  triggered it. The ActiveX case is the one trigger whose expected caller is
  left unstated, because nothing has measured what Excel reports for an
  ActiveX event sink; every other invariant still holds on its rows.

## Adding a test

Copy the nearest existing `.test.ps1` in the suite you are extending. Each file
carries its own `$case = @{...}` data and calls the suite's `_driver.ps1`, so
there is nothing to register and nothing to regenerate — the listing is the
inventory. Files whose names begin with `_` are never discovered as tests.

The manager's own contract — what a test may assume, what it must print, and
how it binds to its Excel — is in
[`StretchXL/StretchXL.md`](../../StretchXL/StretchXL.md).
