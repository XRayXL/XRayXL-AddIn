# APPLICATION.RUN AND SHEET EVENTS ARE NOT WORKSHEET-FUNCTION ENTRIES.
#
# The escape boundary (D92) is "Excel started this frame to compute a cell", found
# by xlfCaller naming a calling cell that differs from the frame beneath. Two kinds
# of entry are NOT cells, and so must NOT read `unhandled`:
#
#   Application.Run "Macro"   xlfCaller is #REF!, so a run macro is not a cell entry.
#   Worksheet_Change          an event has no calling cell either.
#
# An unhandled error in either propagates as VBA rather than becoming a cell's
# #VALUE!, so its frame reads `threw`/`unwound`, never `unhandled`. Excel still pops
# its modal error dialog for these (even under On Error Resume Next -- StretchXL.md);
# the watchdog dismisses them and the test declares them expected.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub RunDriver()
    On Error Resume Next
    Application.Run "RunThrower"
End Sub

Public Sub RunThrower()
    Dim t As Long
    t = 1
    Err.Raise 5, "XRayCase", "raised inside an Application.Run macro"
    t = 2
End Sub

Public Sub ChangeDriver()
    On Error Resume Next
    ThisWorkbook.Worksheets(1).Range("Z1").Value = 1
End Sub
'@

$sheetCode = @'
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$Z$1" Then
        Dim e As Long
        e = 1
        Err.Raise 5, "XRayCase", "raised inside a sheet event"
    End If
End Sub
'@

function ExitsOf($Rows, [string]$Fn) {
    @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}
function OutcomesOf($Rows, [string]$Fn) {
    $e = ExitsOf $Rows $Fn
    if ($e.Count -eq 0) { return '(no row)' }
    return (($e | ForEach-Object { $_.outcome }) -join ',')
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'RunAndEvents' @(
        @{ Kind=1; Name='RunEvt'; Code=$moduleCode }
        @{ Kind='Sheet'; Code=$sheetCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail "the raise slot did not verify on this VBE7: $armLine"
    }

    # An unhandled error in an Application.Run macro or a sheet event pops Excel's
    # modal error dialog even under On Error Resume Next; the watchdog dismisses it
    # and this test declares the dialogs expected (Write-DialogsHandled) so they do
    # not turn its PASS into a FAIL.
    $dlgBefore = @(Get-SessionDialogs).Count
    try { $app.Run($leaf + '!RunDriver') | Out-Null } catch {}
    try { $app.Run($leaf + '!ChangeDriver') | Out-Null } catch {}

    $dlgNew = @(Get-SessionDialogs).Count - $dlgBefore
    if ($dlgNew -gt 0) { Write-DialogsHandled }

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    $rt = @(ExitsOf $rows 'RunThrower')
    Check 'application-run-macro-was-traced' ($rt.Count -ge 1) "RunThrower rows: $($rt.Count)"
    Check 'application-run-error-is-not-a-cell-escape' `
          (@($rt | Where-Object { $_.outcome -eq 'unhandled' }).Count -eq 0) `
          "RunThrower=$(OutcomesOf $rows 'RunThrower') -- must propagate as VBA, not read unhandled"

    $ev = @(ExitsOf $rows 'Worksheet_Change')
    Check 'sheet-event-was-traced' ($ev.Count -ge 1) "Worksheet_Change rows: $($ev.Count)"
    Check 'event-error-is-not-a-cell-escape' `
          (@($ev | Where-Object { $_.outcome -eq 'unhandled' }).Count -eq 0) `
          "Worksheet_Change=$(OutcomesOf $rows 'Worksheet_Change')"

    # The only cell escapes in a session are worksheet functions; there are none here.
    $escapes = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.outcome -eq 'unhandled' })
    Check 'nothing-here-escaped-to-a-cell' ($escapes.Count -eq 0) `
          ("unhandled: " + (@($escapes | ForEach-Object { $_.function }) -join ','))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("Application.Run and a sheet event propagate as VBA, neither a cell escape; $dlgNew dialog(s) dismissed")
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
