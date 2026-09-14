# A REAL Err.Raise NESTED BENEATH A BENIGN OBJECT-MODEL RAISE.
#
# TOUCHING THE OBJECT MODEL RAISES, through the same opcode as Err.Raise, so a
# raise sets the error state only PROVISIONALLY and the frame's own close proves
# it benign. While that state is up, a SECOND raise arriving
# is deduped (the raise opcode fires ~4x per Err.Raise). This asks the sharp
# question that dedupe raises: if the second raise is a GENUINE Err.Raise in a
# NESTED frame, is it still attributed -- or does the benign frame's provisional
# state swallow a real throw?
#
#   P1_Outer  On Error GoTo, catches            -> handled
#   P1_Mid    writes a cell (benign), then calls -> the error unwinds through it
#   P1_Thrower  Err.Raise (the real error)       -> threw  <-- the assertion
#
# If P1_Thrower reads `returned`, a benign cell write suppressed a real throw's
# attribution -- the exact failure the fix must not introduce.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub P1_Outer()
    On Error GoTo Caught
    P1_Mid
    Exit Sub
Caught:
    Dim n As Long
    n = 1
    n = 2
End Sub

Public Sub P1_Mid()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    ws.Range("A1").Value = 1        ' benign object-model raise, in THIS frame
    P1_Thrower
    ws.Range("A2").Value = 9        ' never reached
End Sub

Public Sub P1_Thrower()
    Err.Raise 5, "XRayCase", "a real error raised beneath a benign one"
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P1' @(
        @{ Kind=1; Name='P1Case'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    # an unverified raise slot is a product failure, not a SKIP: every outcome would read `returned`
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail ("the raise slot did not verify on this VBE7, so no outcome " +
                                     "below can be attributed: $armLine")
    }

    $app.Run($leaf + '!P1_Outer') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $t = Outcome $rows 'P1_Thrower'; $m = Outcome $rows 'P1_Mid'; $o = Outcome $rows 'P1_Outer'

    Check 'real-thrower-reads-threw' ($t -eq 'threw') `
          "P1_Thrower='$t' (a benign cell write in P1_Mid must not suppress this real Err.Raise)"
    Check 'catcher-reads-handled' ($o -eq 'handled') "P1_Outer='$o'"
    Write-Output ("outcomes: P1_Thrower=$t  P1_Mid=$m  P1_Outer=$o")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed -- P1_Thrower=$t P1_Mid=$m P1_Outer=$o" }
    Complete-Test -Pass -Detail "real throw beneath a benign raise: P1_Thrower=$t P1_Mid=$m P1_Outer=$o"
}
catch { Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message) }
