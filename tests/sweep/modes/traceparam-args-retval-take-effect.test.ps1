# ARGS and RETVAL change the trace, not just the echo. The same workbook is traced twice, once
# with both flags on and once with both off, and the rows are compared.
#
# A disabled decode changes what is emitted, never what is counted: the frames still appear,
# pair, and name their caller; only `args` and `ret` go empty.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Public Function P_Add(ByVal a As Double, ByVal b As Double) As Double
    P_Add = a + b
End Function
'@
    # B1 is the XLL side: each source has its own switches, and both are flipped.
    New-XRayMacroBook $sx 'ParamEffect' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'A1' = '=P_Add(2,3)'
        'B1' = '=TxB(2,3)'
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $leaf = $book.Leaf

    function Trace-Once([bool]$capture) {
        Invoke-XRayArmedSession $sx -Settings @(@('VBA', 'ARGS', $capture), @('VBA', 'RETVAL', $capture),
                                                 @('XLL', 'ARGS', $capture), @('XLL', 'RETVAL', $capture)) `
            -ArmWait 'VBA tracing: ' -Leaf $leaf -Body {
                Invoke-XRayRecalc $app
                Get-XRayCellText $ws.Range('A1')
            }
    }

    $on  = Trace-Once $true
    $off = Trace-Once $false

    # The workbook must compute the same answer either way -- capture
    # settings are an observation choice, not a change to the calculation.
    Check 'value-unchanged-by-capture' (($on.Result -eq '5') -and ($off.Result -eq '5')) "on='$($on.Result)' off='$($off.Result)'"

    $eOn  = @($on.Rows  | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'P_Add' })
    $xOn  = @($on.Rows  | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA')  -and $_.function -eq 'P_Add' })
    $eOff = @($off.Rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'P_Add' })
    $xOff = @($off.Rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA')  -and $_.function -eq 'P_Add' })

    # Exits too: without them the return-value checks below would be skipped, not failed.
    Check 'traced-both-times' (($eOn.Count -ge 1) -and ($eOff.Count -ge 1) -and ($xOn.Count -ge 1) -and ($xOff.Count -ge 1)) `
          "on=$($eOn.Count)/$($xOn.Count) off=$($eOff.Count)/$($xOff.Count) entries/exits"

    # ---- ON: the columns carry what they always did ------------------------
    if ($eOn.Count) {
        Check 'args-captured-when-on' (((ArgOf $eOn[0] 1) -eq '2') -and ((ArgOf $eOn[0] 2) -eq '3')) "args='$($eOn[0].args)'"
    }
    if ($xOn.Count) {
        Check 'retval-captured-when-on' ($xOn[0].ret -eq '5') "ret='$($xOn[0].ret)' rettype='$($xOn[0].rettype)'"
    }

    # ---- OFF: EMPTY, and only those columns --------------------------------
    if ($eOff.Count) {
        Check 'args-empty-when-off' ([string]::IsNullOrEmpty($eOff[0].args)) "args='$($eOff[0].args)'"
        # The row is still a proper row: it still names its caller and cell,
        # so turning ARGS off costs the arguments and nothing else.
        Check 'caller-survives-args-off' ($eOff[0].caller -eq 'cell') "caller='$($eOff[0].caller)'"
        Check 'cell-survives-args-off' ((Get-CallerCell $eOff[0]) -eq 'A1') "callerref='$($eOff[0].callerref)'"
    }
    if ($xOff.Count) {
        Check 'retval-empty-when-off' ([string]::IsNullOrEmpty($xOff[0].ret)) "ret='$($xOff[0].ret)'"
        Check 'rettype-empty-when-off' ([string]::IsNullOrEmpty($xOff[0].rettype)) "rettype='$($xOff[0].rettype)'"
    }

    # ---- THE XLL SIDE: its own switches, the same rule ----------------------
    $xeOn  = @($on.Rows  | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'XLL' -and $_.function -eq 'TxB' })
    $xeOff = @($off.Rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'XLL' -and $_.function -eq 'TxB' })
    $xxOn  = @($on.Rows  | Where-Object { $_.kind -eq 'exit'  -and $_.source -eq 'XLL' -and $_.function -eq 'TxB' })
    $xxOff = @($off.Rows | Where-Object { $_.kind -eq 'exit'  -and $_.source -eq 'XLL' -and $_.function -eq 'TxB' })
    Check 'xll-traced-both-times' (($xeOn.Count -ge 1) -and ($xeOff.Count -ge 1) -and ($xxOn.Count -ge 1) -and ($xxOff.Count -ge 1)) `
          "on=$($xeOn.Count)/$($xxOn.Count) off=$($xeOff.Count)/$($xxOff.Count) entries/exits"
    if ($xeOn.Count -and $xxOn.Count) {
        Check 'xll-args-captured-when-on' ([string]$xeOn[0].args -ceq 'a1:B=2 a2:B=3') "args='$($xeOn[0].args)'"
        Check 'xll-retval-captured-when-on' ([string]$xxOn[0].ret -eq '23') "ret='$($xxOn[0].ret)'"
    }
    if ($xeOff.Count -and $xxOff.Count) {
        Check 'xll-args-empty-when-off' ([string]::IsNullOrEmpty($xeOff[0].args)) "args='$($xeOff[0].args)'"
        Check 'xll-argcount-kept-when-off' ([string]$xeOff[0].argcount -eq '2') "argcount='$($xeOff[0].argcount)'"
        Check 'xll-retval-empty-when-off' ([string]::IsNullOrEmpty($xxOff[0].ret) -and [string]::IsNullOrEmpty($xxOff[0].rettype)) `
              "ret='$($xxOff[0].ret)' rettype='$($xxOff[0].rettype)'"
    }

    # ---- THE ACCOUNTING IS UNTOUCHED --------------------------------------
    if ($on.Totals -and $off.Totals) {
        $tOn  = ConvertFrom-XRayTotals $on.Totals
        $tOff = ConvertFrom-XRayTotals $off.Totals
        Check 'frames-counted-either-way' ($tOff.framesOpened -eq $tOn.framesOpened) `
              "opened on=$($tOn.framesOpened) off=$($tOff.framesOpened)"
        Check 'no-frame-leak-with-capture-off' ($tOff.framesOpened -eq $tOff.framesClosed) `
              "opened $($tOff.framesOpened), closed $($tOff.framesClosed)"
    }
    else { Check 'totals-present' $false 'no totals line after disarm' }

    $problems = Test-RowInvariants $off.Rows
    Check 'caller-invariants-hold-with-capture-off' ($problems.Count -eq 0) ($problems -join '; ')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'ARGS/RETVAL off empties only those columns; frames, caller and totals unchanged'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
