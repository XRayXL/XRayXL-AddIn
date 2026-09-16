# A Worksheet_Calculate HANDLER THAT TOUCHES THE OBJECT MODEL.
#
# Worksheet_Calculate fires from inside the calc engine itself. The existing
# stress case proves it is traced with a trivial body; this drives the realistic
# shape -- the handler WRITES a cell (a benign object-model raise) and calls a
# helper -- and asks that the write not read as an error and the helper nest
# under it. A volatile UDF makes the event fire deterministically on a full
# recalculation. A non-volatile cell on a clean sheet recomputes nothing, so
# the event never fires and the test sees zero -- hence the volatile UDF.
#
# The handler's own cell write re-dirties the sheet, so the event may fire more
# than once; that is REPORTED, not asserted, because Excel decides whether to
# coalesce those passes and the count is not ours to fix. A module counter caps
# the writes so the recalculation always settles.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public g_wcCount As Long

Public Function Vol() As Double
    Application.Volatile
    Vol = 1
End Function

Public Sub P5_Helper()
    Dim t As Long
    t = 1
End Sub

Public Sub P5_Drive()
    g_wcCount = 0
    Application.CalculateFull
End Sub
'@

$sheetCode = @'
Private Sub Worksheet_Calculate()
    g_wcCount = g_wcCount + 1
    If g_wcCount > 3 Then Exit Sub
    Me.Range("C1").Value = g_wcCount      ' benign object-model write in the handler
    P5_Helper
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P5' @(
        @{ Kind=1; Name='P5Case'; Code=$moduleCode }
        @{ Kind='Sheet'; Code=$sheetCode }
    ) @{
        'A1' = '=Vol()'         # volatile -> a full recalc always fires the event
    }
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!P5_Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows    = Select-BookRows (Read-TraceRows $sx.ProcId) $leaf
    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit'  -and $_.source -eq 'VBA') })
    $names   = @($entries | ForEach-Object { $_.function })
    $calcN   = @($entries | Where-Object { $_.function -eq 'Worksheet_Calculate' }).Count
    $helpN   = @($entries | Where-Object { $_.function -eq 'P5_Helper' }).Count

    # The event fired from inside the calc engine and was traced.
    Check 'calculate-event-traced' ($calcN -ge 1) ("Worksheet_Calculate entries: $calcN  (saw: " + (($names | Select-Object -Unique) -join ',') + ")")
    # Its object-model write did not stop the helper it calls from being traced.
    Check 'helper-under-the-handler-traced' ($helpN -ge 1) "P5_Helper entries: $helpN"

    # The handler wrote a cell; that benign raise must not read as an error, and
    # repeated firing must not corrupt a single frame's outcome.
    $notRet = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'every-frame-returned' ($notRet.Count -eq 0) `
          ("not returned: " + (@($notRet | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    Write-Output ("Worksheet_Calculate fired $calcN time(s); helper $helpN; re-entry is Excel's to coalesce, reported not asserted")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed (calcN=$calcN helpN=$helpN)" }
    Complete-Test -Pass -Detail "Worksheet_Calculate with an object-model write traced clean: fired $calcN, helper $helpN, all returned"
}
catch { Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message) }
