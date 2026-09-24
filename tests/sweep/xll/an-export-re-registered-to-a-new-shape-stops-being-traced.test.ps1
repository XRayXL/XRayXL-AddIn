# An export re-registered while armed to a different shape stops being traced: Excel keeps one
# registration per export, so the planned thunk would forward too few arguments. The log says so,
# and the function keeps working.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    # Re-registered by an earlier run in this Excel, the arm plans the wide shape and nothing
    # changes. A unique name per run does not help: the premise is the narrow binding at arm.
    if (@($app.Evaluate('TxShapeWide(1,2,3,4,5)'))[0] -is [double]) {
        Complete-Test -Skip -Detail 'TxTwoShapes was re-registered before this arm, so the reshape cannot happen'
    }

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ([string]$armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxShapeNarrow(1,2)'
    Invoke-XRayRecalc $app
    $before = [string](Get-XRayCellText $ws.Cells.Item(1, 1))

    $mark2 = Get-LogLength $paths.Log
    $app.Run('TxRegisterReshape') | Out-Null
    [void](Wait-LogLine $paths.Log 'no longer traced' $mark2 20)
    # The per-export line, not the count that follows it.
    $gone = [string](@(Get-Content $paths.Log | Select-Object -Skip $mark2 |
                       Where-Object { $_ -match 'TxTwoShapes.*no longer traced' }) | Select-Object -First 1)

    $ws.Cells.Item(2, 1).Formula = '=TxShapeWide(1,2,3,4,5)'
    Invoke-XRayRecalc $app
    $after = [string](Get-XRayCellText $ws.Cells.Item(2, 1))
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) | Where-Object { $_.proc -eq 'TxTwoShapes' })
    $wide = @($rows | Where-Object { $_.args -match 'a5' })

    Check 'it-was-traced-before-the-reshape' ([bool]($rows.Count -ge 2)) "rows: $($rows.Count)"
    Check 'the-log-says-it-stopped' ([bool]($gone -match 'TxTwoShapes')) "$gone"
    Check 'no-row-claims-the-new-shape' ([bool]($wide.Count -eq 0)) "rows with a5: $($wide.Count)"
    Check 'the-function-still-gets-its-arguments' ([bool]($after -eq '15')) "A2='$after' (A1 was '$before')"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "traced as QBB, dropped on reshape, call still returned $after"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
