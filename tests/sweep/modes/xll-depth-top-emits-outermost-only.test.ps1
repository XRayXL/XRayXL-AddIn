# XLL DEPTH=TOP emits only the outermost call. The echo alone proves nothing: a setter whose
# effect cannot be observed reads exactly like one that works, so the rows are asserted.
#
# TxCallsBack2 re-enters Excel through xlUDF twice, so one formula gives three genuinely nested
# XLL frames. ALL emits all three; TOP emits one.
#
# The counts must not thin: a filter changes what is emitted, never what is counted, so
# XRayXL_GetTraceSummary must still show all three functions as called.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'XllTop' -Cells @{ 'A1' = '=TxCallsBack2(5)' } -Format xlsx
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $leaf = $book.Leaf

    # Rebuild, so a non-volatile UDF is genuinely re-evaluated the second
    # time round rather than served from the last result.
    $pass = {
        Invoke-XRayRecalc $app 'Rebuild'
        [pscustomobject]@{
            Value   = Get-XRayCellText $ws.Range('A1')
            Summary = ConvertFrom-XRayTraceSummary ($app.Run('XRayXL_GetTraceSummary', 'Tx*'))
        }
    }
    $all = Invoke-XRayArmedSession $sx -Settings @(@('XLL', 'DEPTH', 'ALL'), @('VBA', 'DEPTH', 'OFF')) -Body $pass -Leaf $leaf
    $top = Invoke-XRayArmedSession $sx -Settings @(@('XLL', 'DEPTH', 'TOP'), @('VBA', 'DEPTH', 'OFF')) -Body $pass -Leaf $leaf

    # The answer must not depend on how much we chose to write down.
    Check 'value-unchanged-by-depth' ($all.Result.Value -eq $top.Result.Value) "ALL='$($all.Result.Value)' TOP='$($top.Result.Value)'"

    $eAll = @($all.Rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') })
    $eTop = @($top.Rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') })
    $xTop = @($top.Rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL')  })

    Check 'all-emits-the-whole-nest' ($eAll.Count -ge 3) `
          ("ALL entries: " + (@($eAll | ForEach-Object { $_.function }) -join ','))
    Check 'all-states-the-nesting-in-its-columns' (@($eAll | Where-Object { [string]$_.depth -eq '3' }).Count -ge 1) `
          ("ALL depth/parent: " + (@($eAll | ForEach-Object { "$($_.function)=$($_.depth)/$($_.parent)" }) -join ','))
    Check 'top-emits-exactly-one' ($eTop.Count -eq 1) `
          ("TOP entries: " + (@($eTop | ForEach-Object { $_.function }) -join ','))
    if ($eTop.Count -eq 1) {
        Check 'top-emits-the-outermost' ($eTop[0].function -eq 'TxCallsBack2') "got '$($eTop[0].function)'"
        Check 'top-row-is-the-root' ([string]$eTop[0].depth -eq '1' -and [string]$eTop[0].parent -eq '0') `
              "depth='$($eTop[0].depth)' parent='$($eTop[0].parent)'"
    }
    $inner = @($top.Rows | Where-Object { $_.function -eq 'TxB' -or $_.function -eq 'TxCallsBack' })
    Check 'top-drops-the-inner-calls' ($inner.Count -eq 0) `
          ("leaked: " + (@($inner | ForEach-Object { $_.function }) -join ','))

    # A suppressed entry must not leave a dangling exit: the exit writes only
    # when its entry did, so the pairing has to survive the filter.
    Check 'top-exits-pair-with-entries' ($xTop.Count -eq $eTop.Count) `
          "entries=$($eTop.Count) exits=$($xTop.Count)"
    Check 'caller-invariants-hold-under-top' ((Test-RowInvariants $top.Rows).Count -eq 0) `
          ((Test-RowInvariants $top.Rows) -join '; ')

    # ---- COUNTED, NOT EMITTED ---------------------------------------------
    # The two dropped functions must still appear in the summary with calls,
    # because the filter thins the trace and never the accounting.
    $names = @($top.Result.Summary.Rows | ForEach-Object { $_.Function })
    Check 'summary-still-counts-the-dropped-calls' `
          (($names -contains 'TxB') -and ($names -contains 'TxCallsBack')) ($names -join ',')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'TOP emits the outermost XLL call only; inner calls dropped from the trace but still counted'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
