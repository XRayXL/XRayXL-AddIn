# ARMING AGAIN IN THE SAME PROCESS MUST START FROM A CLEAN SLATE.
#
# Every other test in this suite arms once, or twice. A user does not: they
# arm, look, disarm, change a setting, arm again -- for the life of one Excel.
# Nothing in the XLL resets per-thread state between those sessions, and the
# state that survives is the shadow stack the entry/exit pairing depends on:
# `depth` and `frames[].recorded` in src/xll/xlltrace.cpp are __declspec(thread)
# and are written at arm by nobody.
#
# WHAT THAT COSTS, MEASURED. Under DEPTH=TOP an inner call is counted and its
# entry suppressed, and its exit must be suppressed with it -- the exit writes
# only when `f.recorded` says its entry did. A frame left `recorded = true` by
# an EARLIER arming session makes that check answer yes for a frame this
# session never opened, and the trace gains an exit with no entry.
#
# AN EXIT WITH NO ENTRY IS THE ROW SHAPE THIS TOOL MUST NEVER INVENT. An entry
# with no exit reads as a call that hung -- the thing the trace exists to
# report truthfully -- and an exit with no entry asserts a call that never
# happened. Both are silent: every field is well-formed, so nothing but the
# pairing can catch them.
#
# WHY THIS TEST EXISTS RATHER THAN A HARNESS MODE. The defect was first seen
# under -SessionMode ReuseClean, where consecutive tests share one Excel and
# the arm count builds up -- ninth test in its session, reproducible at
# -Parallel 1, invisible at -Parallel 8 and structurally impossible under
# Fresh. Relying on that to catch it again would mean gating on a slow mode
# and hoping the ordering lands. The subject is not session reuse; it is
# re-arming, and a test can do that itself. So this one does, in ONE Excel, in
# the ordinary Fresh gate, deterministically.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# Enough cycles to reach the state the ReuseClean session reached by its ninth
# test; its log showed ~15 arm/disarm pairs before the failure.
$kCycles = 15

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # TxCallsBack2 re-enters Excel through xlUDF twice, so one formula gives
    # genuinely nested XLL frames -- which is what makes the pairing worth
    # asserting at all. AsLoaded, because a formula in a saved workbook is
    # what a user has.
    New-XRayMacroBook $sx 'XllRearm' -Cells @{ 'A1' = '=TxCallsBack2(5)' } -Format xlsx
    $leaf = (Get-XRayMacroBook).Leaf

    # A SECOND BOOK, LEFT OPEN, calling a traced function of its own. This is
    # not decoration: under -SessionMode Reuse a book left behind by an earlier
    # test stays open, and an armed CalculateFullRebuild recalculates it too.
    # Its rows are real and correct; they simply belong to somebody else, and a
    # test that reads them as its own reports the tracer broken when it is not.
    New-XRayMacroBook $sx 'XllRearmOther' -Cells @{ 'A1' = '=TxB(2,3)' } -Format xlsx

    # Rebuild, not Calculate: a non-volatile UDF served from its last
    # result would leave the shadow stack untouched and prove nothing.
    function Invoke-Cycle([string]$depth, [string]$rowsOf = '') {
        Invoke-XRayArmedSession $sx -Settings @(@('XLL', 'DEPTH', $depth), @('VBA', 'DEPTH', 'OFF')) `
            -Body { Invoke-XRayRecalc $app 'Rebuild' } -Leaf $rowsOf
    }

    # ---- dirty the process -------------------------------------------------
    # Alternating depths on purpose: TOP is the mode that suppresses entries,
    # and a frame left behind by a suppressed session is the one that goes on
    # to be written by a later one.
    $armedEvery = $true
    for ($i = 1; $i -le $kCycles; $i++) {
        $cycle = Invoke-Cycle (@('ALL', 'TOP')[$i % 2])
        if ($cycle.ArmLine -notmatch 'armed \d+ of') { $armedEvery = $false }
    }
    Check 'every-cycle-armed' $armedEvery "$kCycles arm/disarm cycles"

    # ---- the session under test -------------------------------------------
    $rows = @((Invoke-Cycle 'TOP' $leaf).Rows)

    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL')  })

    # The DEPTH=TOP contract, restated here so a failure says which half broke.
    Check 'top-still-emits-exactly-one-entry' ($entries.Count -eq 1) `
          ("entries: " + (@($entries | ForEach-Object { $_.function }) -join ','))

    # ---- THE POINT: pairing, by span, in both directions -------------------
    # Counting rows would catch this particular defect, but only by luck --
    # one orphan entry and one orphan exit would cancel. A span is the
    # identity of an activation, so matching them names WHICH row is unpaired.
    $entrySpans = @($entries | ForEach-Object { $_.span })
    $exitSpans  = @($exits   | ForEach-Object { $_.span })

    $orphanExits = @($exits | Where-Object { $entrySpans -notcontains $_.span })
    Check 'no-exit-without-its-entry' ($orphanExits.Count -eq 0) `
          ("orphan exits: " + (@($orphanExits | ForEach-Object { "$($_.function)/span=$($_.span)" }) -join ','))

    $orphanEntries = @($entries | Where-Object { $exitSpans -notcontains $_.span })
    Check 'no-entry-without-its-exit' ($orphanEntries.Count -eq 0) `
          ("orphan entries: " + (@($orphanEntries | ForEach-Object { "$($_.function)/span=$($_.span)" }) -join ','))

    Check 'entries-and-exits-balance' ($entries.Count -eq $exits.Count) `
          "entries=$($entries.Count) exits=$($exits.Count)"

    # A suppressed inner call must be absent from the trace ENTIRELY, not just
    # from its entry rows -- which is how the defect first showed.
    $inner = @($rows | Where-Object { $_.function -eq 'TxB' -or $_.function -eq 'TxCallsBack' })
    Check 'inner-calls-leave-no-rows-at-all' ($inner.Count -eq 0) `
          ("leaked: " + (@($inner | ForEach-Object { "$($_.kind):$($_.function)" }) -join ','))

    Check 'caller-invariants-hold' ((Test-RowInvariants $rows).Count -eq 0) `
          ((Test-RowInvariants $rows) -join '; ')

    # DROPPED EXITS ARE COUNTED AND REPORTED, and must be zero here. This is a
    # different mechanism from the pairing above -- the recorder being
    # re-entered -- and it would produce the same unpaired-entry row, so it is
    # ruled out separately rather than left to be inferred.
    $dropped = @(Get-Content $paths.Log | Select-String 'exit row\(s\) were dropped')
    Check 'no-exits-were-dropped' ($dropped.Count -eq 0) `
          ("disarm warnings: " + (@($dropped | ForEach-Object { $_.Line }) -join ' | '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed after $kCycles re-arms" }
    Complete-Test -Pass -Detail "$kCycles re-arms, then $($entries.Count) entry / $($exits.Count) exit, every span paired"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
