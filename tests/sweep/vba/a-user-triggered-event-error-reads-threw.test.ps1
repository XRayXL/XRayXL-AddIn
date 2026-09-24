# An unhandled error in a user-fired sheet event or an OnTime macro reads `threw`, never
# `unhandled`: with no calling cell it is not a cell escape. A known gap: under-labelled, not wrong.
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

    $dlgBefore = @(Get-SessionDialogs).Count

    # ---- user-triggered event: the harness writes the cell over COM -------------
    # a write from outside VBA leaves no VBA frame beneath, which a macro-driven write cannot
    try { $ws.Range('Z1').Value2 = 1 } catch {}

    # ---- Application.OnTime: Excel calls the macro on its own -------------------
    # Excel's idle loop must run it; its error pops a modal dialog the watchdog dismisses.
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
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
