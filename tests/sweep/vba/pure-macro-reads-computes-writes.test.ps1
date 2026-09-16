# A PURE VBA MACRO, NO UDF AND NO XLL IN SIGHT -- the plainest thing VBA does:
# a Sub is run, it READS values off a worksheet, COMPUTES over them through a
# chain of nested Function calls, and WRITES the result back. No cell ever calls
# a function; nothing crosses into the XLL. This is the control that proves VBA
# tracing is GENERAL -- it observes the interpreter, not merely UDFs reached
# from a recalculation.
#
# WHAT IS ASSERTED, and how each could be wrong rather than a box ticked:
#
#   present   every procedure that ran must have a row -- a Sub that executed
#             but produced none is exactly the invisibility this tool exists
#             to prevent, and a macro entered by Application.Run reaches the
#             interpreter differently from a UDF entered by a recalc
#   named     each is NAMED, not reported as a trailer address
#   paired    entry and exit by SPAN, both directions
#   nested    the call CHAIN is real: XR_SumWeighted runs inside XR_PureMacro,
#             and each XR_Weight runs inside XR_SumWeighted -- a flattened tree
#             would name every procedure yet parent every one at 0
#   looped    XR_Weight is called five times in a For Each; all five must be
#             present and all five parented on the one XR_SumWeighted frame
#   effect    the macro actually WROTE its result -- end to end, not just
#             entered and abandoned
#   outcome   every exit is `returned` -- nothing raised
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub XR_PureMacro()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    Dim total As Long
    total = XR_SumWeighted(ws)                 ' READ + COMPUTE (nested)
    ws.Range("D1").Value = XR_Label(total)     ' WRITE the result (nested)
End Sub

Public Function XR_SumWeighted(ByVal ws As Worksheet) As Long
    Dim c As Range, acc As Long
    For Each c In ws.Range("A1:A5").Cells
        acc = acc + XR_Weight(CLng(c.Value))   ' deeper still, once per cell
    Next c
    XR_SumWeighted = acc
End Function

Public Function XR_Weight(ByVal n As Long) As Long
    XR_Weight = n * 2
End Function

Public Function XR_Label(ByVal n As Long) As String
    XR_Label = "total=" & CStr(n)
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'PureMacro' @(
        @{ Kind=1; Name='PureCase'; Code=$moduleCode }
    ) @{} {
        param($ws)
        for ($i = 1; $i -le 5; $i++) { $ws.Cells.Item($i, 1).Value = $i }   # A1:A5 = 1..5
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!XR_PureMacro') | Out-Null
    $written = [string]$ws.Range('D1').Value2
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows    = Select-BookRows (Read-TraceRows $sx.ProcId) $leaf
    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    $names   = @($entries | ForEach-Object { $_.function })

    # ---- present, and every one of them VBA ------------------------------
    $want = @('XR_PureMacro','XR_SumWeighted','XR_Weight','XR_Label')
    $missing = @($want | Where-Object { $names -notcontains $_ })
    Check 'all-procedures-traced' ($missing.Count -eq 0) `
          ("missing: " + ($missing -join ',') + "   saw: " + ($names -join ','))

    # ---- named, not an address -------------------------------------------
    $unnamed = @($entries | Where-Object { $_.function -match '^0x[0-9A-Fa-f]+$' })
    Check 'no-procedure-reported-as-an-address' ($unnamed.Count -eq 0) `
          ("unnamed: " + (@($unnamed | ForEach-Object { $_.function }) -join ','))

    # ---- paired, both directions -----------------------------------------
    $eSpans = @($entries | ForEach-Object { $_.span })
    $xSpans = @($exits   | ForEach-Object { $_.span })
    $orphanExits   = @($exits   | Where-Object { $eSpans -notcontains $_.span })
    $orphanEntries = @($entries | Where-Object { $xSpans -notcontains $_.span })
    Check 'no-exit-without-its-entry' ($orphanExits.Count -eq 0) `
          ("orphans: " + (@($orphanExits | ForEach-Object { $_.function }) -join ','))
    Check 'no-entry-without-its-exit' ($orphanEntries.Count -eq 0) `
          ("orphans: " + (@($orphanEntries | ForEach-Object { $_.function }) -join ','))

    # ---- the CALL CHAIN, read from the parent spans ----------------------
    # note carries parent=<span>, 0 for a top-level frame. The macro is
    # entered by Application.Run, so XR_PureMacro is the root of THIS chain;
    # XR_SumWeighted must parent on it, and every XR_Weight on XR_SumWeighted.
    function SpanOf([string]$fn) {
        $e = @($entries | Where-Object { $_.function -eq $fn })
        if ($e.Count) { return [string]$e[0].span } else { return '' }
    }
    function ParentOf($entryRow) {
        return [string]$entryRow.parent
    }
    $macroSpan = SpanOf 'XR_PureMacro'
    $sumSpan   = SpanOf 'XR_SumWeighted'

    $sumEntry = @($entries | Where-Object { $_.function -eq 'XR_SumWeighted' })
    Check 'sum-nests-under-the-macro' (($sumEntry.Count -ge 1) -and ((ParentOf $sumEntry[0]) -eq $macroSpan) -and ($macroSpan -ne '')) `
          ("XR_SumWeighted parent=$(if ($sumEntry.Count) { ParentOf $sumEntry[0] } else { '(absent)' }), XR_PureMacro span=$macroSpan")

    $labelEntry = @($entries | Where-Object { $_.function -eq 'XR_Label' })
    Check 'label-nests-under-the-macro' (($labelEntry.Count -ge 1) -and ((ParentOf $labelEntry[0]) -eq $macroSpan) -and ($macroSpan -ne '')) `
          ("XR_Label parent=$(if ($labelEntry.Count) { ParentOf $labelEntry[0] } else { '(absent)' }), XR_PureMacro span=$macroSpan")

    # ---- the LOOP: five weights, all under the one sum frame --------------
    $weights = @($entries | Where-Object { $_.function -eq 'XR_Weight' })
    Check 'weight-called-once-per-cell' ($weights.Count -eq 5) `
          ("XR_Weight entries: $($weights.Count) (A1:A5 is five cells)")
    $misparented = @($weights | Where-Object { (ParentOf $_) -ne $sumSpan })
    Check 'every-weight-nests-under-the-sum' (($weights.Count -ge 1) -and ($misparented.Count -eq 0) -and ($sumSpan -ne '')) `
          ("XR_SumWeighted span=$sumSpan; weight parents: " + (@($weights | ForEach-Object { ParentOf $_ }) -join ','))

    # ---- the macro actually did its job ----------------------------------
    # 2*(1+2+3+4+5) = 30, written as text into D1. If the chain ran but wrote
    # nothing, the trace could look complete while the macro did nothing.
    Check 'macro-wrote-its-result' ($written -eq 'total=30') `
          ("D1='$written' (expected 'total=30')")

    # ---- negative control: nothing raised --------------------------------
    $notReturned = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'every-outcome-is-returned' ($notReturned.Count -eq 0) `
          ("not returned: " + (@($notReturned | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    # ---- caller invariants hold on the whole set -------------------------
    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    Write-Output ''
    Write-Output 'what was traced:'
    foreach ($e in $entries) { Write-Output ("  entry span={0,-4} {1,-16} depth={2} parent={3}" -f $e.span, $e.function, $e.depth, $e.parent) }


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "pure macro traced end to end: $($entries.Count) entries / $($exits.Count) exits, D1='$written'"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
