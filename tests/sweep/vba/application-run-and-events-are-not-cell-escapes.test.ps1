# An error in an Application.Run macro or a sheet event is not a cell escape: neither has a
# calling cell, so neither frame may read `unhandled`.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
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

    # these errors pop a modal dialog even under On Error Resume Next; declared expected below
    $dlgBefore = @(Get-SessionDialogs).Count
    try { $app.Run($leaf + '!RunDriver') | Out-Null } catch {}
    try { $app.Run($leaf + '!ChangeDriver') | Out-Null } catch {}

    $dlgNew = @(Get-SessionDialogs).Count - $dlgBefore
    if ($dlgNew -gt 0) { Write-DialogsHandled }

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    # asserts threw, since "not unhandled" alone would pass a wrong `returned`
    $rt = @(ExitsOf $rows 'RunThrower')
    Check 'application-run-macro-was-traced-once' ($rt.Count -eq 1) "RunThrower rows: $($rt.Count)"
    Check 'application-run-error-reads-threw' `
          (($rt.Count -eq 1) -and ($rt[0].outcome -eq 'threw')) `
          "RunThrower=$(OutcomesOf $rows 'RunThrower') -- must propagate as VBA, reading threw"

    # End on the event's dialog fires no opcode, so disarm flushes the frame as if still
    # running: `returned` with trust `flush` (docs/TraceRowModel.md).
    $ev = @(ExitsOf $rows 'Worksheet_Change')
    Check 'sheet-event-was-traced-once' ($ev.Count -eq 1) "Worksheet_Change rows: $($ev.Count)"
    Check 'event-error-ends-at-the-dialog-and-flushes' `
          (($ev.Count -eq 1) -and ($ev[0].outcome -eq 'returned') -and ($ev[0].trust -eq 'flush')) `
          "Worksheet_Change=$(OutcomesOf $rows 'Worksheet_Change') trust='$(if ($ev.Count) { $ev[0].trust })'"

    # The only cell escapes in a session are worksheet functions; there are none here.
    $escapes = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.outcome -eq 'unhandled' })
    Check 'nothing-here-escaped-to-a-cell' ($escapes.Count -eq 0) `
          ("unhandled: " + (@($escapes | ForEach-Object { $_.function }) -join ','))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("Application.Run and a sheet event propagate as VBA, neither a cell escape; $dlgNew dialog(s) dismissed")
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
