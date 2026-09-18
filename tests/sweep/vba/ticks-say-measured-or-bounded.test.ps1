# Is `ticks` a measurement, or an upper bound? The row says which.
#
# CloseFrame stamps a row with the time it runs, and only one of its callers runs when the
# activation actually ends:
#
#    exit opcode        fires at the end of the activation      -> a MEASUREMENT
#    stack-pointer      fires at the NEXT statement, at a        -> an UPPER BOUND
#      backstop         higher rsp
#    flush at disarm    fires at the end of the session          -> an UPPER BOUND
#
# A fully unhandled unwind fires no exit opcodes (docs/TraceRowModel.md), so every frame in a
# blown-up call chain is closed by the backstop. The number is kept: the frame ended at or
# before that instant.
#
# The unhandled error is raised from a cell, not Application.Run, which would open Excel's modal
# VBA error dialog. In a formula it becomes #VALUE! and unwinds silently.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
' No handler anywhere: the unwind reaches the cell and becomes #VALUE!.
Public Function T_Boom(ByVal n As Long) As Long
    T_Boom = T_BoomInner(n)
End Function

Public Function T_BoomInner(ByVal n As Long) As Long
    Dim z As Long
    z = 1
    Err.Raise 5, "XRayCase", "unhandled, on purpose"
    T_BoomInner = z
End Function

' A clean chain, so `exit` is asserted against something in the same session.
Public Function T_Fine(ByVal n As Long) As Long
    T_Fine = T_FineInner(n) + 1
End Function
Public Function T_FineInner(ByVal n As Long) As Long
    T_FineInner = n * 2
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    function ClosedBy($row) {
        if ($row.trust) { return [string]$row.trust }
        return '(none)'
    }

    # AsLoaded: the formulas are in the saved workbook, which is what a user has.
    New-XRayMacroBook $sx 'Ticks' @(
        @{ Kind=1; Name='TickCase'; Code=$moduleCode }
    ) @{
        'A1' = '=T_Fine(10)'
        'A2' = '=T_Boom(1)'
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    Invoke-XRayRecalc $app 'Rebuild'
    $a1 = Get-XRayCellText $ws.Range('A1')
    $a2 = Get-XRayCellText $ws.Range('A2')
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $totals = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark 20
    $rows   = @(Read-TraceRows $sx.ProcId)
    $exits  = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })

    # The unwind has to have actually happened, or everything below is vacuous.
    Check 'the-unhandled-call-really-blew-up' ($a2 -eq '#VALUE!') "A2 held '$a2' (expected #VALUE!)"
    Check 'the-clean-call-really-worked'      ($a1 -eq '21')      "A1 held '$a1' (expected 21)"

    # ---- every exit row declares which it is ------------------------------
    $vocab = @('exit','backstop','flush')
    $bad = @($exits | Where-Object { (ClosedBy $_) -notin $vocab })
    Check 'every-exit-row-says-how-it-closed' ($bad.Count -eq 0) `
          ("rows without a valid closed=: {0} of {1}" -f $bad.Count, $exits.Count)

    $boomRows  = @($exits | Where-Object { $_.function -like 'T_Boom*' })
    $fineRows  = @($exits | Where-Object { $_.function -like 'T_Fine*' })

    # ---- the clean chain is MEASURED --------------------------------------
    $fineNotExit = @($fineRows | Where-Object { (ClosedBy $_) -ne 'exit' })
    Check 'clean-calls-are-measured' (($fineRows.Count -ge 2) -and ($fineNotExit.Count -eq 0)) `
          ("T_Fine rows: " + (@($fineRows | ForEach-Object { "$($_.function)=$(ClosedBy $_)" }) -join ','))

    # ---- the unwound chain is BOUNDED -------------------------------------
    # This is the assertion the whole change exists for: a fully unhandled
    # unwind fires no exit opcodes, so its frames MUST be marked as bounded.
    $boomBounded = @($boomRows | Where-Object { (ClosedBy $_) -in @('backstop','flush') })
    Check 'unhandled-unwind-is-marked-bounded' `
          (($boomRows.Count -ge 1) -and ($boomBounded.Count -eq $boomRows.Count)) `
          ("T_Boom rows: " + (@($boomRows | ForEach-Object { "$($_.function)=$(ClosedBy $_)" }) -join ','))

    # ---- and it is not marking everything ---------------------------------
    # A change that stamped every row `backstop` would satisfy the case above.
    Check 'not-everything-is-bounded' ($fineRows.Count -gt 0) `
          ("measured rows: {0}, bounded rows: {1}" -f `
            @($exits | Where-Object { (ClosedBy $_) -eq 'exit' }).Count,
            @($exits | Where-Object { (ClosedBy $_) -ne 'exit' }).Count)

    # ---- the totals and the rows must agree -------------------------------
    # Two independent measures of the same thing: the counter is incremented in
    # CloseFrame, the token is written in EmitRow. If they disagree, one of them
    # is lying and the disagreement says so.
    $tBackstop = 0; $tFlush = 0
    if ($totals -match 'closedByBackstop=(\d+)') { $tBackstop = [int]$Matches[1] }
    if ($totals -match 'closedByFlush=(\d+)')    { $tFlush    = [int]$Matches[1] }
    $rBackstop = @($exits | Where-Object { (ClosedBy $_) -eq 'backstop' }).Count
    $rFlush    = @($exits | Where-Object { (ClosedBy $_) -eq 'flush' }).Count
    Check 'totals-agree-with-the-rows' (($tBackstop -eq $rBackstop) -and ($tFlush -eq $rFlush)) `
          "totals backstop=$tBackstop flush=$tFlush; rows backstop=$rBackstop flush=$rFlush"

    # ---- the outcome column still says what happened ----------------------
    $threw = @($exits | Where-Object { $_.outcome -eq 'threw' })
    Check 'the-thrower-is-still-named' ($threw.Count -ge 1) `
          ("threw: " + (@($threw | ForEach-Object { $_.function }) -join ','))

    # An uncaught error must not contaminate what comes after. Nothing catches the A2 error, so
    # the chain ends when the shadow stack empties, and the in-flight flag must clear with it.
    # Asserted on the clean rows in A1, which would otherwise read `unwound`.
    $fineWrong = @($fineRows | Where-Object { $_.outcome -ne 'returned' })
    Check 'an-uncaught-error-does-not-leak-into-later-calls' ($fineWrong.Count -eq 0) `
          ("clean rows: " + (@($fineRows | ForEach-Object { "$($_.function): ticks=$($_.ticks) trust=$($_.trust)" }) -join ' | '))

    # ...and the escape is COUNTED, so `threw` exceeding `handled` is explained
    # rather than left as an unexplained shortfall.
    $escaped = 0
    if ($totals -match 'errEscaped=(\d+)') { $escaped = [int]$Matches[1] }
    Check 'the-escape-is-counted' ($escaped -ge 1) `
          "errEscaped=$escaped (the A2 error left VBA uncaught)"

    Write-Output ''
    Write-Output 'exit rows:'
    foreach ($x in $exits) { Write-Output ("  {0,-14} ticks={1,-10} trust={2}" -f $x.function, $x.ticks, $x.trust) }


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("measured={0} bounded={1} across a clean chain and an unhandled unwind" -f `
        @($exits | Where-Object { (ClosedBy $_) -eq 'exit' }).Count, $boomBounded.Count)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
