# A FUNCTION REGISTERED AFTER ARMING BY AN ADD-IN THAT ASKS FOR NO RESULT.
#
# Excel12(xlfRegister, 0, ...) is how the SDK's own sample registers, so no
# register id comes back to the watch. The registration still happened, and the
# function must be hooked rather than dismissed as one Excel refused.
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
    $app.Run('TxRegisterNoResult') | Out-Null
    $lateLine = Wait-LogLine $paths.Log 'late arm:' $mark2 20
    Write-XRayObservation 'late-arm' "$lateLine"

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxLateNoResult(3)'
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) | Where-Object { $_.function -eq 'TxLateNoResult' })

    Check 'excel-accepted-the-registration' ($value -eq '21') "A1='$value'"
    Check 'the-late-arm-did-not-call-it-refused' ($lateLine -notmatch 'refused') "$lateLine"
    Check 'the-function-was-traced' ($rows.Count -ge 2) "rows naming TxLateNoResult: $($rows.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "traced $($rows.Count) row(s) of a function registered with no result"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
