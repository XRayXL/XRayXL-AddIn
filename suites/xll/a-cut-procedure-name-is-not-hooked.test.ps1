# A PROCEDURE NAME LONGER THAN THE REGISTRATION CAPTURE.
#
# The watch keeps 127 characters of a procedure name. Cut, a long name can spell
# a shorter export exactly, and hooking that one would trace an unrelated
# function under the long one's name. TracedAddin registers such a name after
# arming, then TxCallsPrefix calls the shorter export the cut name spells.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }
    $watchLine = Wait-LogLine $paths.Log 'register watch: (installed|NOT installed|DISABLED)' $mark 15
    if ($watchLine -notmatch 'installed on MdCallBack12') { Complete-Test -Fail -Detail "watch: $watchLine" }

    $mark2 = Get-LogLength $paths.Log
    $app.Run('TxRegisterLongName') | Out-Null
    # A declined capture logs nothing, so this waits for a hook that should not happen, briefly.
    $applied = Wait-LogLine $paths.Log 'register watch: applied' $mark2 5
    Write-XRayObservation 'watch-after-long-name' "$applied"
    $registered = [string]$app.Evaluate('TxLongNamed(1)')

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxCallsPrefix(2)'
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $misnamed = @($rows | Where-Object { $_.function -eq 'TxLongNamed' })
    $caller = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'TxCallsPrefix' })

    Check 'excel-registered-the-long-name' ($registered -eq '1.75') "TxLongNamed(1)='$registered'"
    Check 'the-caller-answered' ($value -eq '2.5') "A1='$value'"
    Check 'the-caller-was-traced' ($caller.Count -ge 1) "TxCallsPrefix entries: $($caller.Count)"
    Check 'no-row-carries-the-cut-name' ($misnamed.Count -eq 0) `
          "rows named TxLongNamed: $($misnamed.Count), though TxLongNamed was never called while armed"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'a cut procedure name hooked nothing'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
