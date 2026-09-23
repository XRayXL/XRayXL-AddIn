# AN ERROR RAISED ON THE FIRST LINE OF A NESTED CALLEE.
#
# A procedure's frame opens at its first statement, and a raise on that line
# fires before it, while the CALLER is still the top of the shadow stack. The
# raise has to wait for the callee's frame at any depth, or the caller is
# reported as the thrower and the real thrower as a clean return.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub RF_Outer()
    Dim n As Long
    On Error GoTo Caught
    n = 1
    RF_Middle
    n = 2
    Exit Sub
Caught:
    n = 3
    n = 4
End Sub

Public Sub RF_Middle()
    Dim m As Long
    m = 1
    RF_Thrower
    m = 2
End Sub

Public Sub RF_Thrower()
    Err.Raise 5, "XRayCase", "a deliberate error"
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'FirstLine' @(
        @{ Kind=1; Name='FirstLineCase'; Code=$moduleCode }
    )
    $leaf = (Get-XRayMacroBook).Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!RF_Outer') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $t = Outcome $rows 'RF_Thrower'; $m = Outcome $rows 'RF_Middle'; $o = Outcome $rows 'RF_Outer'

    Check 'the-callee-that-raised-says-threw' ($t -eq 'threw') "RF_Thrower outcome='$t'"
    Check 'its-caller-says-unwound' ($m -eq 'unwound') "RF_Middle outcome='$m'"
    Check 'the-catcher-says-handled' ($o -eq 'handled') "RF_Outer outcome='$o'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "chain: RF_Thrower=$t -> RF_Middle=$m -> RF_Outer=$o"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
