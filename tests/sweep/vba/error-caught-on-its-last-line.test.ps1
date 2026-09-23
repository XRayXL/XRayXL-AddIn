# AN ERROR CAUGHT BY A FRAME WHOSE NEXT STATEMENT IS ITS OWN END.
#
#     Function LL_Catch() As Long
#         On Error Resume Next
#         LL_Catch = LL_Boom()      ' raises; what runs next is End Function
#     End Function
#
# `handled` goes to a frame that predates the raise and runs again. Here the
# only thing the catcher runs is its epilogue, so that has to count. Otherwise
# the catcher reads `returned`, and the catch is charged to its caller or, at
# the top, counted as an error escaping VBA.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Function LL_Boom() As Long
    Dim t As Long
    t = 1
    Err.Raise 5, "XRayCase", "a deliberate error"
End Function

Public Function LL_Catch() As Long
    On Error Resume Next
    LL_Catch = LL_Boom()
End Function

' A caller with more to do, where a misplaced catch would land.
Public Sub LL_Top()
    Dim v As Long
    v = LL_Catch()
    v = v + 1
End Sub

' The same catcher as the outermost frame, where a miss reads as an escape.
Public Function LL_CatchAtTop() As Long
    On Error Resume Next
    LL_CatchAtTop = LL_Boom()
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'LastLine' @(
        @{ Kind=1; Name='LastLineCase'; Code=$moduleCode }
    )
    $leaf = (Get-XRayMacroBook).Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!LL_Top') | Out-Null
    $app.Run($leaf + '!LL_CatchAtTop') | Out-Null
    $mark2 = Get-LogLength $paths.Log
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $totals = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2

    $rows = @(Read-TraceRows $sx.ProcId)
    $boom = Outcome $rows 'LL_Boom'; $catch = Outcome $rows 'LL_Catch'
    $top  = Outcome $rows 'LL_Top';  $atTop = Outcome $rows 'LL_CatchAtTop'
    $escaped = if ($totals -match 'errEscaped=(\d+)') { [int]$Matches[1] } else { -1 }
    $threw   = if ($totals -match '\bthrew=(\d+)')    { [int]$Matches[1] } else { -1 }
    $handled = if ($totals -match '\bhandled=(\d+)')  { [int]$Matches[1] } else { -1 }

    Check 'the-thrower-says-threw' ($boom -eq 'threw') "LL_Boom outcome='$boom'"
    Check 'a-catcher-that-only-returns-says-handled' ($catch -eq 'handled') "LL_Catch outcome='$catch'"
    Check 'its-caller-is-not-given-the-catch' ($top -eq 'returned') "LL_Top outcome='$top'"
    Check 'the-same-catcher-at-the-top-says-handled' ($atTop -eq 'handled') "LL_CatchAtTop outcome='$atTop'"
    Check 'nothing-escaped-vba' ($escaped -eq 0) "errEscaped=$escaped"
    Check 'every-throw-was-handled' (($threw -eq 2) -and ($handled -eq 2)) "threw=$threw handled=$handled"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "Boom=$boom Catch=$catch Top=$top CatchAtTop=$atTop errEscaped=$escaped"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
