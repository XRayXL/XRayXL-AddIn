# Calls with NO CALLING CELL: Application.Run of a command, Application.Run
# and Application.Evaluate of a UDF, and genuine re-entry through Evaluate.
# The unambiguous assertion is that the macro really ran (its own counter
# moved); what the trace attributes to a cell-less call is reported into the
# captured log.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arming did not succeed: $armLine" }

    $macroBefore = [double]$app.Evaluate('TxMacroCalls()')
    foreach ($probe in @(
        @{ N = 'run-TxMacro';            S = { $app.Run('TxMacro') } }
        @{ N = 'run-TxB-2-3';            S = { $app.Run('TxB', 2, 3) } }
        @{ N = 'evaluate-TxB';           S = { $app.Evaluate('TxB(2,3)') } }
        @{ N = 'evaluate-TxCallsBack';   S = { $app.Evaluate('TxCallsBack(5)') } }
    )) {
        try {
            $r = & $probe.S
            Write-Output ("{0} -> {1}" -f $probe.N, $r)
        } catch {
            Write-Output ("{0} -> EXCEPTION: {1}" -f $probe.N, $_.Exception.Message)
        }
    }
    $macroAfter = [double]$app.Evaluate('TxMacroCalls()')
    $moved = ($macroAfter -gt $macroBefore)
    Write-TestCase -Name 'application-run-invoked-the-command' -Pass:$moved -Fail:(-not $moved) `
                   -Detail "count $macroBefore -> $macroAfter"

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $rows = Read-TraceRows $sx.ProcId
    # CALLERS, ASSERTED. Run and Evaluate reach the function with no calling
    # cell, and the contract says that decodes as caller='none' -- never as an
    # invented cell. Measured: this test produces exactly 'cell' and
    # kind 'none', nothing else.
    $inv = @(Test-RowInvariants $rows)
    Write-TestCase -Name 'caller-invariants-hold' -Pass:($inv.Count -eq 0) -Fail:($inv.Count -ne 0) -Detail ($inv -join '; ')
    $noCell = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -and -not (Get-CallerCell $_) })
    $wrong = @($noCell | Where-Object { $_.caller -ne 'none' })
    $ok = ($noCell.Count -gt 0 -and $wrong.Count -eq 0)
    Write-TestCase -Name 'cell-less-calls-say-none-ref' -Pass:$ok -Fail:(-not $ok) `
                   -Detail ("{0} cell-less entries, {1} off-vocabulary" -f $noCell.Count, $wrong.Count)
    foreach ($e in ($noCell | Select-Object -First 8)) {
        Write-Output ("  {0} caller={1} args=[{2}]" -f $e.function, $e.caller, $e.args)
    }

    if ($moved) { Complete-Test -Pass -Detail "macro count $macroBefore->$macroAfter, $($rows.Count) trace rows" }
    else { Complete-Test -Fail -Detail 'Application.Run did not actually invoke the command' }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
