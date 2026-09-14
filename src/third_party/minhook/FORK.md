# Local changes to the vendored MinHook

This copy is **upstream MinHook** (<https://github.com/TsudaKageyu/minhook>) at
commit `d94c64d` (2026-06-13). Apart from line endings and byte-order marks,
every file is identical to that commit except for the local change below. The
same files are unchanged upstream back to `1e9ad1e` (2025-11-03); upstream's
next commit, `8af6b4a` (a Hacker Disassembler Engine update), is not included.
m417z's fork carries the same commits on its `master` branch.

The queue API — `MH_QueueEnableHook` / `MH_QueueDisableHook` / `MH_ApplyQueued`,
upstream since 2013 — is what this project depends on: batching turned arming 45
functions from 3,496 ms into 79 ms, because MinHook suspends every thread in the
process on each enable.

## The one local change: how threads are enumerated for the freeze

`MH_SetThreadFreezeMethod(MH_FREEZE_METHOD_FAST_UNDOCUMENTED)`, which upstream
does not have. The API, its names and the `NtGetNextThread` approach come from
Michael Maltsev's (m417z) commit `17b9da3` (2021-09-10, "Improve thread
suspend/resume and allow to choose an alternative method"), which is on m417z's
`multihook` branch and on the `threading_patch` branch of
<https://github.com/TFORevive/minhook>. What is here is a smaller port of it.

**Why.** `Freeze()` finds the threads it must suspend with
`CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, 0)`. That flag is documented to
snapshot **every thread on the machine** regardless of the process id argument,
and MinHook then filters for its own. Measured inside Excel:

    thread snapshot: 60,091us for 6,564 threads system-wide, 81 ours

That 60 ms *was* the cost of a batch patch — the patching itself is about 15 us
per hook. So a whole-process freeze cost 60 ms of thread enumeration to do 0.7 ms
of work.

**What changed.** `EnumerateThreadsFast()` uses `NtGetNextThread`, which walks
one process. It fills the **same `LPDWORD` array of thread ids**, so `Freeze()`
and `Unfreeze()` are otherwise untouched and still `OpenThread()` by id.

**What was deliberately not taken.** Commit `17b9da3` also changes
`FROZEN_THREADS` to hold `HANDLE`s and suspends during enumeration. That is
faster still — it saves two `OpenThread` calls per thread — but it rewrites
`Freeze`, `Unfreeze` and the original enumerator too. Against a vendored
dependency, a diff small enough to read in full is worth more than saving ~81
handle opens. Its third mode, `MH_FREEZE_METHOD_NONE_UNSAFE`, is not ported:
nothing here wants unsuspended patching.

**Failure behaviour.** `NtGetNextThread` is undocumented (Vista+). If it cannot
be resolved, `MH_SetThreadFreezeMethod` returns an error and the original method
stays in force: slower, otherwise the same.

## Re-applying after a MinHook update

Four marked hunks (`LOCAL ADDITION`) and one unmarked line:
  * `include/MinHook.h` — the enum and the declaration
  * `src/hook.c` — the `NtGetNextThread` typedef and globals, near the top
  * `src/hook.c` — `EnumerateThreadsFast()`, just before `Freeze()`
  * `src/hook.c` — the enumerator choice inside `Freeze()` (unmarked)
  * `src/hook.c` — `MH_SetThreadFreezeMethod()`, at the end

If upstream adopts a process-scoped enumeration, delete these and take theirs.
m417z's `multihook` branch already has one, but it also reworks hook management
(about 2,300 lines differ from this copy), so moving to it is a larger change
than re-applying these hunks.
