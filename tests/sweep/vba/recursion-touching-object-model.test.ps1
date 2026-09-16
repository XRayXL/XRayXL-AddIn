# RECURSION THAT WRITES A CELL AT EVERY LEVEL.
#
# The object-model fix narrows the `unwound` mark to frames OUTER than the raise
# (span < errSpan). Recursion is where that comparison is most exposed: spans
# rise with depth, so a benign raise deep in the stack sets a HIGH errSpan while
# many shallower frames are still live. This drives a five-deep recursion whose
# every frame writes a cell (a benign raise) and asks that NONE of them read as
# an error -- the fix must hold when the span it compares against moves.
#
# Nothing here throws, so every frame must read `returned`, the chain must reach
# depth 5, and the parents must chain P4_Rec(n) under P4_Rec(n+1).
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub P4_Drive()
    Dim r As Long
    r = P4_Rec(4)
End Sub

Public Function P4_Rec(ByVal n As Long) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    ws.Range("D" & (n + 1)).Value = n     ' a benign object-model write at each level
    If n > 0 Then
        P4_Rec = P4_Rec(n - 1) + n
    Else
        P4_Rec = 0
    End If
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P4' @(
        @{ Kind=1; Name='P4Case'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!P4_Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows    = Select-BookRows (Read-TraceRows $sx.ProcId) $leaf
    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit'  -and $_.source -eq 'VBA') })
    $recEx   = @($exits | Where-Object { $_.function -eq 'P4_Rec' })

    Check 'five-recursion-levels-traced' ($recEx.Count -eq 5) "P4_Rec exits: $($recEx.Count) (expected 5, n=4..0)"

    $notRet = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'every-frame-returned' ($notRet.Count -eq 0) `
          ("frames not returned: " + (@($notRet | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    # Depth: P4_Drive(1) -> P4_Rec(2) -> ... -> P4_Rec(6). Read the deepest depth= note.
    $maxDepth = 0
    foreach ($e in $entries) { if ($e.depth) { $d = [int]$e.depth; if ($d -gt $maxDepth) { $maxDepth = $d } } }
    Check 'recursion-reached-depth-6' ($maxDepth -ge 6) "deepest depth=$maxDepth (P4_Drive + five P4_Rec)"

    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed (recEx=$($recEx.Count) depth=$maxDepth)" }
    Complete-Test -Pass -Detail "recursion touching the object model: $($recEx.Count) levels, depth $maxDepth, all returned"
}
catch { Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message) }
