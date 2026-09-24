# A function registered after arming with Excel12(xlfRegister, 0, ...), as the SDK's sample does,
# gets no register id back, but must still be hooked rather than dismissed as refused.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    # TxRegisterNoResult picks a fresh Excel name every call and returns which one, so a
    # reused session cannot find it already registered and skip the late path.

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }
    $watchLine = Wait-LogLine $paths.Log 'register watch: (installed|NOT installed|DISABLED)' $mark 15
    if ($watchLine -notmatch 'installed on MdCallBack12') { Complete-Test -Fail -Detail "watch: $watchLine" }

    $mark2 = Get-LogLength $paths.Log
    $which = [int]$app.Run('TxRegisterNoResult')
    $lateName = "TxLateNoResult$which"
    $lateLine = Wait-LogLine $paths.Log 'late arm:' $mark2 20
    Write-XRayObservation 'late-arm' "$lateLine"

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = "=$lateName(3)"
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) | Where-Object { $_.function -eq $lateName })

    Check 'excel-accepted-the-registration' ($value -eq '21') "A1='$value'"
    Check 'the-late-arm-hooked-it' ($lateLine -match 'late arm:.*hooked [1-9]') "$lateLine"
    Check 'the-late-arm-did-not-call-it-refused' ($lateLine -notmatch 'refused') "$lateLine"
    Check 'the-function-was-traced' ($rows.Count -ge 2) "rows naming $lateName`: $($rows.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "traced $($rows.Count) row(s) of a function registered with no result"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
