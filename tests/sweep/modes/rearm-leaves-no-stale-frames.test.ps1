# Arming again in the same process must start from a clean slate: the per-thread shadow stack
# survives between sessions unless arming resets it, and under DEPTH=TOP a frame left `recorded`
# would give an exit with no entry, every field well-formed.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# Enough arm/disarm cycles to reach the state a long reused session reaches.
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

    # A second book, left open, calling a traced function of its own: an armed
    # CalculateFullRebuild recalculates it too. Its rows are real and belong to somebody else.
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

    # ---- the point: pairing, by span, in both directions -------------------
    # Counting rows could let one orphan entry and one orphan exit cancel; matching spans names
    # which row is unpaired.
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

    # A suppressed inner call must be absent from the trace entirely, not just from its entry rows.
    $inner = @($rows | Where-Object { $_.function -eq 'TxB' -or $_.function -eq 'TxCallsBack' })
    Check 'inner-calls-leave-no-rows-at-all' ($inner.Count -eq 0) `
          ("leaked: " + (@($inner | ForEach-Object { "$($_.kind):$($_.function)" }) -join ','))

    Check 'caller-invariants-hold' ((Test-RowInvariants $rows).Count -eq 0) `
          ((Test-RowInvariants $rows) -join '; ')

    # Dropped exits are counted and reported, and must be zero: a re-entered recorder would produce
    # the same unpaired entry, so it is ruled out separately.
    $dropped = @(Get-Content $paths.Log | Select-String 'exit row\(s\) were dropped')
    Check 'no-exits-were-dropped' ($dropped.Count -eq 0) `
          ("disarm warnings: " + (@($dropped | ForEach-Object { $_.Line }) -join ' | '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed after $kCycles re-arms" }
    Complete-Test -Pass -Detail "$kCycles re-arms, then $($entries.Count) entry / $($exits.Count) exit, every span paired"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
