# AN On Error GoTo HANDLER THAT ITSELF TOUCHES THE OBJECT MODEL.
#
# The real cleanup pattern: catch, then write a status cell. The handler now
# fires its OWN benign object-model raise WHILE the tracer is resolving the
# error episode the thrower started. This asks whether the handler's benign
# raise collides with the `handled` detection.
#
#   P3_Thrower  Err.Raise                         -> threw
#   P3_Outer    On Error GoTo; handler writes cells -> handled
#
# If P3_Outer reads anything but `handled`, the handler's cell write disturbed
# the catch it was in the middle of.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub P3_Outer()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    On Error GoTo Caught
    P3_Thrower
    Exit Sub
Caught:
    ws.Range("A1").Value = "cleanup"   ' benign object-model write IN the handler
    ws.Range("A2").Value = "done"
End Sub

Public Sub P3_Thrower()
    Err.Raise 5, "XRayCase", "a real error whose handler does object-model cleanup"
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P3' @(
        @{ Kind=1; Name='P3Case'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!P3_Outer') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $t = Outcome $rows 'P3_Thrower'; $o = Outcome $rows 'P3_Outer'

    Check 'thrower-reads-threw' ($t -eq 'threw') "P3_Thrower='$t'"
    Check 'catcher-reads-handled-despite-om-cleanup' ($o -eq 'handled') `
          "P3_Outer='$o' (its handler wrote two cells -- benign raises must not disturb the catch)"
    Write-Output ("outcomes: P3_Thrower=$t  P3_Outer=$o")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed -- P3_Thrower=$t P3_Outer=$o" }
    Complete-Test -Pass -Detail "handler with object-model cleanup: P3_Thrower=$t P3_Outer=$o"
}
catch { Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' })) }
