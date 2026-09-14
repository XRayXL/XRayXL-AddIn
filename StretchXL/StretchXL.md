# StretchXL

A generic harness for testing things that run inside Excel — add-ins, XLLs,
workbooks, automation — built around one question: **does an Excel session
start, do what it is asked, and shut down cleanly?**

Nothing in the core knows what it is testing. Each test is its own process,
handed an Excel the manager owns, and answers on stdout. The manager keeps
what is hard: lifecycles, deadlines, real exit codes, dumps taken before
anything is killed, and the measured close-down that comes free with every
test ever written.

## Contents

**Using it** — [Quick start](#quick-start) · [Layout](#layout) · [Parameters](#parameters) · [Output](#output) · [Reading a hang](#reading-a-hang)

**Writing tests** — [The test contract](#the-test-contract) · [Suite configuration](#suite-configuration) · [Session lifetime](#session-lifetime) · [Order is a variable, and always reproducible](#order-is-a-variable-and-always-reproducible)

**How it behaves** — [What the manager provides](#what-the-manager-provides) · [Deadlines, and leaks attributed](#deadlines-and-leaks-attributed) · [The floor is the control, forever](#the-floor-is-the-control-forever) · [The crash toolkit travels with the tool](#the-crash-toolkit-travels-with-the-tool)

**Why it is built this way** — [Measured behaviours this is built around](#measured-behaviours-this-is-built-around) · [Testing an unhandled VBA error needs a cell, not Application.Run](#testing-an-unhandled-vba-error-needs-a-cell-not-applicationrun) · [Why the harness releases every COM reference](#why-the-harness-releases-every-com-reference) · [What stays out of the core](#what-stays-out-of-the-core) · [Rules](#rules) · [Safety](#safety)

---

## Quick start

```powershell
# everything under a tests root
.\StretchXL\StretchXL.ps1 -Parallel 8 -Path .\suites -OutDir <any folder>

# the floor: no -Path, so no tests -- the control that proves the harness clean
.\StretchXL\StretchXL.ps1 -Parallel 8 -OutDir <any folder>

# a subset -- the folder IS the selection
.\StretchXL\StretchXL.ps1 -Parallel 4 -Path .\suites\stress -OutDir <any folder>

# the order-dependency soak: reused dirty sessions, shuffled, replayable by seed
.\StretchXL\StretchXL.ps1 -Parallel 8 -Runs 20 -SessionMode Reuse -RandomOrder -Path .\suites -OutDir <any folder>
```

Requires 64-bit Excel and an interactive desktop session. `-Parallel` and
`-OutDir` are mandatory; there are no hidden temp locations.

## Layout

**Infrastructure and tests live in separate locations.** The tests folder is
given per run (`-Path`), and the folder *is* the selection mechanism: point at
the root to run everything, at a subfolder to run that group. Discovery is recursion from wherever you pointed; there is no registry
to maintain.

```
StretchXL\
  StretchXL.ps1          the manager: discovery, workers, lifecycle, reporting
  TestKit.ps1            the helper a PowerShell test dot-sources for the contract
  _common.ps1            what the manager and the kit have to agree on: the
                         suite.psd1 resolver, the refusals it implies, and
                         kill-by-pid-and-name. Dot-sourced by both
  evidence\              the crash toolkit: hangdump.ps1, hangwhere.py, dumpstack.py
  StretchXL.md           this file

<tests root>\            anywhere; handed to the manager as -Path
  <area>\                one folder per suite (nesting allowed)
    suite.psd1           declarative session set-up: add-ins to load, settings
                         (optional; nearest wins per key, unset keys inherited
                         from ancestor configs; paths resolve relative to the
                         file, %VAR% environment references expand)
    <name>.test.ps1      one file per test; only *.test.ps1 files are tests
```

## The test contract

**The contract is environment in, stdout out — language-agnostic on purpose.**
Each test runs as its own process. Under the manager, a worker owns the Excel
session end to end — it starts it, ledgers it, hands it out, resets it, and
closes it, measured — and describes it to the test in environment variables;
the test binds to exactly that instance by window handle
(`AccessibleObjectFromWindow` / `OBJID_NATIVEOM` — never `GetActiveObject`,
which binds "some Excel" rather than this one), does its one thing, and
reports a verdict. A live COM interface marshals across processes by design,
but does not survive PowerShell's job serialization — binding by hwnd is what
lets tests live in their own processes without losing the identity guarantee.

```
manager -> test (environment):
  STRETCH_SESSION_HWND   the Excel to bind to; absent => standalone, start your own
  STRETCH_SESSION_PID    to verify the binding landed on the right process
  STRETCH_SESSION_DIR    this session's own directory -- what the add-in under
                         test wrote and what the test writes, kept together
  STRETCH_TEST_ID        the test's own name, so its scratch is named after
                         itself and not after a pid Windows will re-issue
  STRETCH_WORKDIR        the run's work root, above the session directories
  STRETCH_LEDGER         where any pid the test starts is recorded before use
  STRETCH_DIALOG_LOG     the watchdog's log of dialogs it has dismissed, for a
                         test that means to judge them itself

test -> manager (stdout + exit code):
  STRETCH verdict=PASS|FAIL|SKIP detail="..."      one per test, mandatory
  STRETCH case=<name> verdict=PASS|FAIL detail=... optional, before the verdict,
                                                   for tests that check many
                                                   cases -- counted and reported
                                                   per case by the manager
  anything else on stdout is captured log; no verdict line, or a nonzero
  exit, is an ERROR
```

Tests run under `powershell.exe` (5.1) — the edition every COM behaviour in
this file was measured on. PowerShell tests
implement all of this by dot-sourcing `TestKit.ps1` and calling
`Connect-TestExcel` / `Complete-Test`; nothing requires a test to be
PowerShell. Run standalone — `.\name.test.ps1` from any prompt, no manager,
no environment — the kit starts an Excel of its own and closes it afterwards,
because **the party who started a session closes it**: that one rule is why
the measured close-down always belongs to the manager when the manager is
running things, and why a test never times, kills, or closes what it did not
start.

## What the manager provides

Around every test: the Excel, ledgered before use; the close (Quit or
simulated X click), measured; two deadlines enforced from outside; hang dumps
taken before anything is killed; real exit codes from its own process handle;
warm-up; parallelism; the dialog watchdog; and the streamed one-line-per-run
reporting with counted outcomes.

**Two axes of outcome, always both reported.** What the test concluded
(`PASS` / `FAIL` / `SKIP`, with detail) and what the session did
(`CLEAN` / `ALIVE` / `CRASH` / `HANG` / `EXITED` / `FAILED`, with `shutdown=`;
`ALIVE` is a reused session that survived this test, whose close is measured
later on its own `session-close` line). They are
independent facts: a test can pass and Excel then crash on the way down — that
is a result, not two half-results — and the close-down axis comes free with
every test ever written, which is the point of owning the lifecycle centrally.

## Deadlines, and leaks attributed

A test body that never returns is a
`TIMEOUT` of the test — its process is killed from outside, cleanly, because
it is a process; an Excel that will not exit afterwards is a `HANG` of the
session. Releasing what you acquire is still the contract: a test process
that exits normally takes its stragglers with it, but one that had to be
killed leaves its references to DCOM's rundown — minutes, not milliseconds —
so a close measured after a killed test says as much about the kill as about
Excel, and the report says which it was.

## Session lifetime

Fresh-per-test is the default and
proves isolation. Reuse mode hands a sequence of tests one long-lived
session. Reusing a session reuses the PROCESS, never its workbooks: every
book a test opens is closed when that test finishes, because a book left
open is recalculated by the next test's armed CalculateFull and pollutes its
trace (a soak found exactly that). **The one book spared is identified by
NAME** — the manager's baseline, the one its bind window lives on, recorded at
connect time; everything else is closed whether or not it was ever saved.
Deciding it by SHAPE instead (a saved book has a `Path`, the baseline does not)
silently spares any book a test created with `Workbooks.Add()` and never saved
— and `Quit` hides that, because Quit closes every book at once. **An X click
closes one window**, so the process stays alive holding the extra book and runs
out the close deadline as a HANG. What persists is the process dirt --
Application settings, loaded XLLs, VBA and in-XLL state -- because "fails only
in a dirty session" is how Excel fails on real machines, and a fresh-only
framework can never see it. Reuse comes in two
flavours, chosen per run: **dirty** (`Reuse`) — the realism arm, each test
inherits whatever the last one left behind — and **cleaned** (`ReuseClean`) —
where the worker itself puts the session back to baseline between tests: every
workbook except the baseline closed unsaved, and `DisplayAlerts`,
`Calculation`, `ScreenUpdating`, `EnableEvents` and `StatusBar` restored. The
reset is best-effort and says so: document and settings state can be restored;
a loaded XLL, a COM add-in's state, anything cached in a module cannot —
undoing those takes a fresh process, which is what fresh mode is for.

**A session is only reused between tests whose set-up matches**: the same
`RegisterXll` list, `RegisterXllSettleSeconds` and `SessionEnvironment`. The
first test with a different set-up closes the session (its `session-close`
line) and opens a new one, so no test inherits another suite's add-ins or
environment. Tests in different folders with the same set-up share a session,
so a shuffled run does not cycle Excel at every folder boundary.
`RequireNotOlderThan` is not part of the match: it is checked before any Excel
starts and changes nothing inside one.

Reuse is paid for honestly: every result line carries the session and the
test's position in it, `shutdown=` is measured once per session rather than
per test and reported on its own `session-close` line, and a session that dies
during a test is that test's outcome — `CRASH` or `EXITED`, read from the
manager's own process handle — after which the next item simply opens a
replacement session and the run continues.

## Suite configuration

A `suite.psd1` beside the tests declares the session every test in that folder
should start from. Resolution climbs **above** the `-Path` root for as long as
each level carries one, so pointing at a subfolder inherits the suite's config
exactly as running the whole suite would — a subset run that quietly skipped
the XLL registration is a measured failure, not a hypothetical. The nearest
file wins per key; unset keys are inherited; paths resolve relative to the file
and `%VAR%` environment references expand.

| Key | What it does |
|---|---|
| `RegisterXll` | The add-ins to load into each session before its tests run |
| `RegisterXllSettleSeconds` | How long to let registration settle |
| `TestTimeoutSeconds` | Raise the body deadline for a slow suite (overrides are printed) |
| `ModulesOfInterest` | Modules `hangwhere.py` should report on when triaging a dump |
| `SessionEnvironment` | Name/value pairs set in the worker's environment *before* Excel starts — a COM-created Excel inherits its creator's environment, and so does the test's PowerShell. `{SessionDir}` expands to the per-session directory (also exported as `STRETCH_SESSION_DIR`). Cleared before the next session, so nothing leaks into the floor |
| `RequireNotOlderThan` | A map of file → the file or directory it derives from, or a list of them when it is built from several. A directory means its newest source file (`.cpp`, `.h`, `.c`, `.hpp`, `.asm`, `.inc`, `.def`). Checked **before any Excel starts**; a violation refuses the run rather than reporting anything |

**`RequireNotOlderThan` is the one that has earned its keep twice.** A skipped
install step otherwise lets a run certify the *previous* binary in green, which
is worse than red because nothing about it looks wrong — measured, as a 215/215
sweep and a commit saying "full suite green", both against an add-in nine
commits behind. Naming a *directory* as the reference means "no older than the
newest source file under it", which closes the other end: a build that fails
leaves the previous binary with its old timestamp, the deploy copies that same
binary, the two timestamps agree, and the run certifies a build that never
happened.

**Set-up is declarative, and there is no set-up script.** What a suite can ask
for is the table above, applied by the worker once per session before any test
sees it: the environment exported, the XLLs registered, the settle waited. A
failure there is its own verdict — `SETUP-FAILED` — so the test is never
spawned against a half-made session and a config mistake cannot masquerade as a
test failure. Set-up also defines **baseline**: `ReuseClean` restores a session
to set-up-complete, not to a virgin Excel, because the XLLs stay registered.

**A test run standalone gets the same set-up**, from the same resolver
(`_common.ps1`): the kit walks the suite chain above the test's own folder,
refuses the same missing or stale artifacts, exports the same environment
before it starts its Excel, and registers the same XLLs. "It passes standalone"
and "it passes in the sweep" have to be answers to the same question.

## Order is a variable, and always reproducible

`-RandomOrder` shuffles
the tests — combined with dirty reuse, that is the soak that flushes out
order-dependent failures, where a test only breaks after some other test has
run. The shuffle is driven by a seed printed in the header and settable with
`-Seed`, because a random order that found a failure and cannot be replayed
is half a result.

## The floor is the control, forever

With no test given, the harness runs
start / add a workbook / close — the least Excel can be asked to do — and that
is what keeps the infrastructure provably clean: 0 hangs, 0 crashes, shutdown
2.3–2.6s, serial and parallel, on the measured baseline. Any hang the floor
cannot reproduce and a test can was caused by what the test does.

## What stays out of the core

Anything product-specific — loading a particular XLL,
arming a tracer, asserting on a trace file — belongs in test scripts or a
suite's shared helper, configured through parameters and context, never
hard-coded in the harness. Tools that orchestrate several Excels of their own
or need a debugger attached stay separate instruments; the harness does not
try to be them.

## The crash toolkit travels with the tool

`evidence\` holds the harness's own copies of `hangdump.ps1` (dump a stuck
process while it is still stuck), `hangwhere.py` (what every thread is parked
on) and `dumpstack.py` (the minidump parser both lean on), with the
modules-of-interest a parameter rather than a hard-coded product name. They are
shutdown-investigation scripts: they belong to the harness rather than to
whatever it is testing, and they ship with it.

## Parameters

| Parameter | Required | Why it exists |
|---|---|---|
| `-Parallel <int>` | **yes** | Number of Excel sessions running concurrently. Mandatory because concurrency changes timing and timing is what is being measured. It **divides** the work; it does not multiply it. |
| `-Runs <int>` | no (default 1) | How many times each selected test runs (at the floor: total floor iterations). Workers partition the work, built pass-by-pass so an unshuffled soak still interleaves every test each pass. |
| `-Path <string>` | no | The tests root; every `*.test.ps1` under it runs, and pointing at a subfolder runs the subset. Omitted means the floor: start Excel, add one empty workbook, close. |
| `-OutDir <string>` | **yes** | Where `results-<timestamp>.jsonl` (and dumps, absent `-DumpDir`) land. Required on every run so there are no hidden temp locations. |
| `-DumpDir <string>` | no | Separate home for minidumps, when dumps and results should not share a folder. |
| `-SessionMode <enum>` | no (Fresh) | `Fresh` gives each test its own Excel. `Reuse` hands consecutive tests one long-lived session, keeping the process dirt; `ReuseClean` also resets settings and workbooks between them. Both need `-Path` — the floor is the fresh-session control by definition. |
| `-CloseTimeoutSeconds <int>` | no (default 70) | When to call a session a `HANG`. A classifier, not a measurement — the real close-to-exit time is reported on every session as `shutdown=`. A deadline shorter than the slowest clean shutdown manufactures hangs instead of measuring them. |
| `-TestTimeoutSeconds <int>` | no (default 120) | When to call a test body a `TIMEOUT`: both the Excel and the wedged test process are dumped, then the test process is killed. Distinct from the close deadline because they are different failures. |
| `-RandomOrder` / `-Seed <int>` | no | Shuffle the work; seed printed and recorded, `-Seed` replays. With dirty reuse, this is the order-dependency soak. |
| `-CloseWithX` | no | Close by simulated X click instead of `Application.Quit()`. See below. |
| `-Warmup` | no | One unmeasured iteration per worker before the counted runs, through the same machinery (ledger, deadline, dump-on-hang). Exists because the first X-close per worker process costs ~60s; the warm-ups run in parallel, so the invocation pays ~60s once and the measured data loses its known first-run mode. Reported as a `run=warmup` line — a warm-up that hangs is still loud — but never counted: not in the tally, the rates, the shutdown distribution, or the expected-runs check. |
| `-DumpAfterSeconds <int>` | no (default 0 = off) | If Excel is still alive this long after the close, write a minidump **while it is still stuck**, then keep waiting. A `HANG` is always dumped before it is killed; this additionally catches slow-but-eventually-clean shutdowns, which a dump taken only at the deadline can never see. |
| `-FullDump` | no | Full-memory dumps instead of stacks-and-handles (about 20x the size); only needed when something must be read out of the heap. |
| `-NoDump` | no | Never dump, even on `HANG`. For unattended soaks where disk is the constraint. |
| `-Cleanup` | — | Kill any Excel left behind by an interrupted run, using the pid ledgers, and exit. Exists because workers are child processes that outlive a killed parent. |

### `-CloseWithX`: the two ways Excel closes

`Application.Quit()` over COM and a user clicking the X button run different
teardown code inside Excel — the interactive path tears down the UI, asks
about unsaved work and notifies add-ins in an order the automation path does
not repeat. An add-in that hangs close-down for a user can look healthy under
`Quit`, and the reverse, so a regression suite needs both arms.

The simulation is faithful: the window is made visible and posted
`WM_SYSCOMMAND` / `SC_CLOSE`, the exact message the X button posts. It is
*posted*, never sent — `SendMessage` would block the worker inside the very
hang it is measuring. Visibility is part of the simulation, not a second
variable: the X button only exists on a visible Excel, and a visible Excel
does not exit merely because its last COM reference is released, so in this
arm the close message has to do the work and a clean exit proves the
interactive path. The window is shown minimized **without activation**
(`SW_SHOWMINNOACTIVE` before `Visible`) — measured: the foreground window
does not change and the interactive path still engages, so an X-arm soak runs
while you work in other applications.

**Known behaviour of this arm** (measured on this machine): the *first*
X-closed Excel of a worker process shuts down in ~60.5s — reproducibly, and
cleanly; in pure floor runs every later one from the same worker takes ~2.5s.
Warmth does not transfer between processes, but the waits run in parallel
without contending, so any number of workers warming at once costs ~60s of
wall clock in total, once. `-Warmup` pays that cost outside the measurement.
A dump taken mid-slowness shows the main thread blocked in a wait inside
interactive teardown with the Click-to-Run virtualization layer
(`AppVIsvSubsystems64.dll`) and `mso*` on its stack — a 60-second timeout in
something Office calls on the interactive exit path, not a defect in the
harness or the workload. The Quit arm has no such mode. Keep
`-CloseTimeoutSeconds` above it, and expect the X arm's shutdown distribution
to be bimodal by design. **Not fully understood:** with test processes binding
to sessions in between, the ~60s cost has recurred on a later session of an
already-warmed worker, so the per-client-process model taken from the pure
floor is incomplete.

## Output

One line per result, to stdout, as it completes — and the same record, with
the captured test output embedded, as a JSON line in the results file:

```
worker=1 run=1 pid=23184 outcome=CLEAN shutdown=2.3s elapsed=3.3s t=3.4s
worker=1 test=proof/binding.test.ps1 run=1 pid=34612 verdict=PASS outcome=CLEAN shutdown=2.3s test_s=1.1s elapsed=4.5s t=4.5s cases=3/3
worker=2 test=calc/soak.test.ps1 run=5 pid=11996 verdict=PASS outcome=HANG shutdown=70s test_s=8.2s elapsed=79s t=356.1s waited=70s window='' dump='...\EXCEL.11996...hang.dmp'
```

| Field / outcome | Why it exists |
|---|---|
| `verdict=` + `outcome=` | The two axes, both always reported: what the test concluded, and what the session did. `verdict=PASS outcome=HANG` is a real line this harness has produced — the test was right and Excel still would not close, which is a result, not a contradiction. |
| `shutdown=` | The measured close-to-exit time. The distribution is the result; the hang count is only where `-CloseTimeoutSeconds` happened to fall. A regression that moves p50 from 2.3s to 8s never trips a hang counter. |
| `test_s=` | The test body's own duration, separate from session start and close. |
| `cases=` | Passed/total `STRETCH case=` lines from a multi-case test. |
| `t=` | Wall clock since the worker started. A failure that begins at a fixed iteration *count* and one that begins at a fixed elapsed *time* are different mechanisms, and without both numbers they are indistinguishable. |
| `ALIVE` | A reused session survived this test and is being handed to the next one. Benign: its close is measured once, later, on that session's own `session-close` line. |
| `EXITED` | The process went away but the exit code could not be read. Never folded into `CLEAN`: "could not read the code" and "exited zero" are different facts. |
| `FAILED` | The iteration could not even establish which Excel was ours; flagged loudly because that Excel never reached the ledger. |
| `dump=` | Where the evidence went. |

The summary prints counts per outcome, crash and hang rates **with their
denominator**, and min/p50/p90/max of the shutdown times — quantiles, because
a mean over a bimodal distribution describes a shutdown that never happens.
The header states the total, the worker count and the split (e.g.
`split : 5 + 5 + 5 + 5`) so what was actually run is never inferred.

### Why exit codes come from `GetExitCodeProcess`

`System.Diagnostics.Process.ExitCode` reads back as `$null` — silently, no
exception — for any process the object did not itself start, which is every
Excel this harness watches. So each worker opens its own handle
(`OpenProcess`, SYNCHRONIZE + QUERY_LIMITED_INFORMATION) **before** the close,
waits on it (`WaitForSingleObject` — a real wait, immune to pid recycling),
and reads the code from it. A crash fast enough to beat a late-opened handle
would otherwise be reported `CLEAN`, which is the worst possible failure for
an instrument whose headline output is a crash rate.

**And why the exit code, not the add-in's own fault counters, is the
authority.** `STATUS_STACK_BUFFER_OVERRUN` (`0xC0000409`) is a `__fastfail`: it
terminates the process outright and **cannot be caught by SEH**. So an add-in
whose every hook is `__try`-wrapped can die this way with its circuit breaker
untripped and every fault counter reading zero — the guards are not broken, they
are never reached. An instrument that asked the add-in "did you fault?" would
report a clean run over a corpse. The exit code is outside the dead process and
cannot be fooled that way.

Three unrelated events used to print as one "the process is gone" message, and
pooling them corrupted every death-rate comparison made before the distinction
existed:

| Exit code | What happened |
|---|---|
| `0` | the process quit itself — a clean shutdown, no fault |
| non-zero | a fault, or a kill (`Process.Kill()` reports `0xFFFFFFFF`) |
| unreadable | the handle was opened too late, or not retained — reported `EXITED`, never `CLEAN` |

## Measured behaviours this is built around

Excel and COM facts established while building the harness. Each one silently
broke something before it was understood, and each is now load-bearing.

- **`Process.ExitCode` is silently `$null`** for any process the object did not
  itself start — no exception, just null. So a session's exit code is read from
  our own handle with `GetExitCodeProcess`; a *test* process is one we started,
  so its `.ExitCode` is valid.
- **Every dot in a COM chain creates an intermediate wrapper.** Unreleased ones
  keep Excel alive as a message-pump zombie for the full deadline, even after
  the test process has exited. Hence `Clear-ComStragglers`, and tests holding
  and releasing each link.
- **In P/Invoke from PowerShell 5.1, `$null` for a string marshals as an EMPTY
  string** (`[NullString]::Value` is the real null), and the literal
  `0xFFFFFFF0` parses as int32 −16. Both broke the `OBJID_NATIVEOM` bind
  silently until they were measured.
- **A dirty workbook turns the close into a hidden save prompt** that waits the
  full deadline. Before a test session's close the worker marks every open
  workbook `Saved = $true` — the harness never saves anything, so the close
  means discard. Only the save prompt is neutralised; every other dialog still
  appears and still fails loudly.
- **The X arm shows its window minimized without activation**
  (`SW_SHOWMINNOACTIVE`): the foreground window does not move and the
  interactive close path still engages, so an X-arm soak can run while you work.

## Testing an unhandled VBA error needs a cell, not `Application.Run`

An error that reaches the top of a VBA call chain under `Application.Run`
raises Excel's **modal VBA error dialog** and waits for a human, so no script
gets past it — the first attempt at this lost its Excel window and the run was
reported as a crash. The *same* error raised in a **UDF called from a formula**
becomes `#VALUE!` in the cell and unwinds silently: same unwind inside the
interpreter, no dialog, and automatable.

So a test that needs a fully unhandled unwind drives it from a **formula**.
This is the sharpest instance of Rule 6 in [Rules](#rules) — the dialog *is* the evidence,
and here the fix is to pick a trigger that never raises one.

## Why the harness releases every COM reference

An out-of-process COM server is reference counted. A close — `Quit()` over COM
or the user's X — on an Excel that an automation client still holds a
reference to leaves the process alive **by design**: Excel is waiting for the
release. Dropping the variables (`$x = $null`) releases nothing; the .NET
wrappers keep their counts until a garbage collection eventually runs.

So the harness releases every reference explicitly, in reverse order of
acquisition, immediately after the close, against an Excel still there to
receive it. A test client that leaks references manufactures the exact hang it
exists to detect — an Excel in Task Manager, obediently pumping messages
forever — and no result from such a client says anything about Excel or an
add-in.

## Reading a hang

`python StretchXL\evidence\hangwhere.py <dump>` prints where every thread is
parked. Two signatures matter:

- **A refcount zombie** — the client's fault: every thread in a plain wait,
  the main thread idling in a `win32u` message pump, window gone. Excel has
  processed the close and is waiting, correctly, for a COM reference the
  client never released.
- **A real close-down hang** — worth investigating: a thread blocked on a
  lock, an RPC, or `DllMain`, with the suspect module on its stack.

## Rules

1. **All results go to stdout, streamed.** One line per run, as it completes.
   Never `Write-Host`: it writes to the information stream, which a pipeline or
   a scraper never sees. `2>&1 | Out-String` does **not** capture it either, so
   a harness that scrapes its own child gets an empty string and every regex
   over it matches nothing; capture with `*>&1`. An afternoon of paired
   comparisons once returned `0/N` for that reason alone, and the tell was that
   *every* comparison, across unrelated hypotheses, returned double-zero:
   **uniform nulls are a broken instrument, not a low base rate.**
2. **Progress is visible while it runs.** A wedged run must look different from
   a slow one *at the time*, not in hindsight. The parent prints a `...`
   heartbeat whenever nothing has been reported for 10 seconds.
3. **Every outcome is counted, including the boring ones.** An instrument that
   cannot tell "nothing happened" from "we never looked" is not an instrument.
4. **A hang is a result, not cleanup.** It is recorded — and dumped — before
   anything is killed. A killed process leaves no evidence; a hung one is the
   best witness there is, every thread parked exactly where it stuck.
5. **Deadlines are enforced by the waiter**, never by the thing being waited
   on.
6. **An unexpected dialog is dismissed, recorded, and failed — never silent.**
   A dialog is Excel asking a question the test did not expect, and the
   question is the evidence, so a per-worker watchdog finds it by class and by
   the session's pid, records its text, and clicks it away (`End` before
   `OK`/`Cancel`, so a VBA runtime error does not leave the interpreter mid-
   procedure or open the editor). A `PASS` with an unacknowledged dialog is
   turned into a `FAIL`; a test that means to judge the dialogs itself reads
   `STRETCH_DIALOG_LOG` and says `STRETCH dialogs=handled`, which waives the
   conversion and nothing else — the count still lands on the record.
7. **Kill only processes this harness started, by process id**, double-checked
   against the image name because pids are recycled. Other people's Excel
   sessions share this machine, with unsaved work in them.
8. **Never modify a user file.** Workbooks are created in the harness's own
   directory or opened read-only.

## Safety

Every Excel's pid is written to a ledger *before* the process is used, so an
interrupted worker's Excel can always be found. Each invocation has its own
ledger file, so two harness runs at once cannot kill each other's Excels; the
parent clears its own ledger in a `finally`, and `-Cleanup` sweeps all ledgers
for the case where even the `finally` never ran. Kills are by recorded pid,
double-checked against the image name. The harness writes only inside
`-OutDir`, `-DumpDir` and `%TEMP%\StretchXL`, and changes no machine
configuration. `-Cleanup` sweeps standalone TestKit ledgers too: they are
named to match, so one pattern finds both.
