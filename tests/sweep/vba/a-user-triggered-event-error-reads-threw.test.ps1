# AN ERROR IN A VBA ACTIVATION EXCEL STARTED WITH NO CALLING CELL READS `threw`.
#
# The escape boundary (D92) is a worksheet-function entry, found by xlfCaller naming
# a calling cell. Some activations Excel starts have no calling cell: a sheet event
# fired by a user edit, and an Application.OnTime macro (scheduled Now+1s, given a
# 10s idle window to fire). xlfCaller answers `#REF!` for both, so they are not cell escapes. With no VBA frame beneath, an unhandled error in
# one reaches the bottom of the shadow stack and reads `threw` (never `unhandled`,
# never `handled`), and the escape is counted. This is the one documented gap: it is
# under-labelled, not wrong.
#
# The event is fired FROM OUTSIDE VBA -- the harness writes the cell over COM -- so
# the handler runs as a top-level VBA activation with nothing of VBA beneath it, which
# is the user-triggered shape a macro-driven write cannot reproduce.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$sheetCode = @'
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$Z$1" Then
        Dim e As Long
        e = 1
        Err.Raise 5, "XRayCase", "raised in a user-triggered event"
    End If
End Sub
'@

$moduleCode = @'
Public Sub OnTimeThrower()
    Dim t As Long
    t = 1
    Err.Raise 5, "XRayCase", "raised in an OnTime macro"
End Sub
'@

function ExitsOf($Rows, [string]$Fn) {
    @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}
function RowsOf($Rows, [string]$Fn, [string]$Kind) {
    @($Rows | Where-Object { $_.kind -eq $Kind -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'UserEvent' @(
        @{ Kind=1; Name='EvtMod';  Code=$moduleCode }
        @{ Kind='Sheet'; Code=$sheetCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf
    $ws = $book.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail "the raise slot did not verify on this VBE7: $armLine"
    }

    $dlgBefore = @(Get-SessionDialogs).Count

    # ---- USER-TRIGGERED EVENT: the harness writes the cell over COM -------------
    # This fires Worksheet_Change with no VBA frame beneath; the handler raises.
    try { $ws.Range('Z1').Value2 = 1 } catch {}

    # ---- APPLICATION.ONTIME: Excel calls the macro on its own -------------------
    # Schedule for one second out and leave Excel idle for ten -- long enough that its
    # idle loop runs the macro. Like the event it has no calling cell, so its unhandled
    # error reads `threw`, and it pops the modal dialog the watchdog dismisses.
    $app.OnTime((Get-Date).AddSeconds(1), "$leaf!OnTimeThrower") | Out-Null
    Start-Sleep -Seconds 10

    $dlgNew = @(Get-SessionDialogs).Count - $dlgBefore
    if ($dlgNew -gt 0) { Write-DialogsHandled }

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    # ---- the event: threw, at the top, its own cell not named -------------------
    $evEntry = @(RowsOf $rows 'Worksheet_Change' 'entry')
    $evExit  = @(ExitsOf $rows 'Worksheet_Change')
    Check 'the-event-was-traced' ($evExit.Count -ge 1) "Worksheet_Change exit rows: $($evExit.Count)"
    Check 'the-event-reads-threw' `
          (($evExit.Count -ge 1) -and (@($evExit | Where-Object { $_.outcome -ne 'threw' }).Count -eq 0)) `
          "Worksheet_Change outcome(s): $((@($evExit | ForEach-Object { $_.outcome })) -join ',')"
    Check 'the-event-is-not-a-cell-escape' `
          (@($evExit | Where-Object { $_.outcome -eq 'unhandled' }).Count -eq 0) `
          "an event has no calling cell, so it must not read unhandled"
    Check 'the-event-has-no-calling-cell' `
          (($evEntry.Count -ge 1) -and (@($evEntry | Where-Object { $_.caller -eq 'cell' }).Count -eq 0)) `
          "callers: $((@($evEntry | ForEach-Object { $_.caller })) -join ',')"
    Check 'the-event-is-a-top-level-activation' `
          (($evEntry.Count -ge 1) -and (@($evEntry | Where-Object { $_.parent -ne '0' }).Count -eq 0)) `
          "parents: $((@($evEntry | ForEach-Object { $_.parent })) -join ',')"

    # ---- OnTime: same category as an event, and asserted --------------------------
    $ot = @(ExitsOf $rows 'OnTimeThrower')
    Check 'ontime-macro-fired-and-was-traced' ($ot.Count -ge 1) "OnTimeThrower exit rows: $($ot.Count)"
    Check 'ontime-macro-reads-threw' `
          (($ot.Count -ge 1) -and (@($ot | Where-Object { $_.outcome -ne 'threw' }).Count -eq 0)) `
          "OnTimeThrower outcome(s): $((@($ot | ForEach-Object { $_.outcome })) -join ',')"
    Check 'ontime-macro-is-not-a-cell-escape' `
          (@($ot | Where-Object { $_.outcome -eq 'unhandled' }).Count -eq 0) `
          "OnTime has no calling cell, so it must not read unhandled"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("a user-triggered event and an OnTime macro both read threw, not unhandled; $dlgNew dialog(s) dismissed")
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
