# When Application.Caller is a range (a CSE array formula), the trace must say so, through both the
# XLL and VBA hooks, while a single-cell caller in the same run stays a cell. TxRetQArray puts a
# multi-cell caller and a multi-cell return on the same span.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx  = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Public Function CA_Scalar() As Double
    CA_Scalar = 42#
End Function
'@
    New-XRayMacroBook $sx 'CallerArray' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{} {
        param($ws)
        # CSE, not a dynamic array: a dynamic array spills from one anchor cell, which is its caller.
        $ws.Range('B2:D4').FormulaArray = '=TxB(2,3)'          # XLL scalar, 9 cells
        $ws.Range('F2:F5').FormulaArray = '=CA_Scalar()'       # VBA scalar, 4 cells
        # A 2x2 return over a 2x2 range: multi-cell caller and multi-cell return.
        $ws.Range('H2:I3').FormulaArray = '=TxRetQArray(7)'    # XLL array,  4 cells
        # The same function in one cell: what a function returns says nothing about where it was called from.
        $ws.Range('A4').Formula         = '=TxRetQArray(9)'    # XLL array,  1 cell
        $ws.Range('A1').Formula         = '=TxB(4,5)'          # XLL scalar, 1 cell (control)
        $ws.Range('A2').Formula         = '=CA_Scalar()'       # VBA scalar, 1 cell (control)
    }

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm refused: $pressed" }

    Invoke-XRayRecalc $app 'Rebuild'
    [void](Invoke-XRayDisarm $sx)

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $entries = @($rows | Where-Object { $_.kind -eq 'entry' })
    Check 'something-was-traced' ($entries.Count -gt 0) "$($entries.Count) entry row(s)"
    if ($entries.Count -eq 0) { Complete-Test -Fail -Detail 'no entry rows at all' }

    # Every caller here is a real cell on a real sheet, whatever its shape.
    $callers = @($entries | Where-Object { $_.caller -eq 'cell' })
    Check 'callers-are-cells' ($callers.Count -gt 0) `
        ("caller values seen: " + (($entries | ForEach-Object { $_.caller } | Sort-Object -Unique) -join ', '))

    function CellsFor([string]$fn) {
        @($callers | Where-Object { $_.function -ieq $fn } | ForEach-Object { Get-CallerCell $_ } | Sort-Object -Unique)
    }

    # ---- the XLL side ------------------------------------------------------
    $txb = CellsFor 'TxB'
    Check 'xll-array-caller-names-the-whole-range' ([bool]($txb -contains 'B2:D4')) `
        ("TxB caller cells: " + ($txb -join ' '))
    Check 'xll-single-cell-caller-is-not-a-range' ([bool]($txb -contains 'A1')) `
        ("TxB caller cells: " + ($txb -join ' '))

    # ---- the VBA side, through the same decoder ----------------------------
    $vba = CellsFor 'CA_Scalar'
    Check 'vba-array-caller-names-the-whole-range' ([bool]($vba -contains 'F2:F5')) `
        ("CA_Scalar caller cells: " + ($vba -join ' '))
    Check 'vba-single-cell-caller-is-not-a-range' ([bool]($vba -contains 'A2')) `
        ("CA_Scalar caller cells: " + ($vba -join ' '))

    # ---- a function that returns an array ----------------------------------
    $ret = CellsFor 'TxRetQArray'
    Check 'array-returning-xll-array-caller-names-the-range' ([bool]($ret -contains 'H2:I3')) `
        ("TxRetQArray caller cells: " + ($ret -join ' '))
    Check 'array-returning-xll-single-cell-caller-stays-one-cell' ([bool]($ret -contains 'A4')) `
        ("TxRetQArray caller cells: " + ($ret -join ' '))

    # The return is still decoded as an array in both placements: the caller's shape must not change it.
    $retRows = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.function -ieq 'TxRetQArray' })
    $retTypes = @($retRows | ForEach-Object { $_.rettype } | Sort-Object -Unique)
    Check 'array-return-still-decoded-whatever-the-caller-shape' `
        ($retRows.Count -ge 2 -and ($retTypes | Where-Object { $_ }).Count -ge 1) `
        ("{0} exit row(s), rettype(s): {1}" -f $retRows.Count, (($retTypes | Where-Object { $_ }) -join ' '))

    # No caller may be the bare first cell of an array range: that is what truncation produces, and
    # no formula here occupies B2, F2 or H2 alone.
    $truncated = @($callers | Where-Object { (Get-CallerCell $_) -in @('B2','F2','H2') })
    Check 'no-caller-truncated-to-its-first-cell' ($truncated.Count -eq 0) `
        ("truncated rows: " + (($truncated | ForEach-Object { "$($_.function)=$(Get-CallerCell $_)" } | Sort-Object -Unique) -join ' '))

    # The sheet travels with the cell whatever the shape -- the reader asserts
    # this globally, but a range is the case most likely to break it.
    $sheetless = @($callers | Where-Object { -not (Get-CallerSheet $_) })
    Check 'range-callers-still-name-their-sheet' ($sheetless.Count -eq 0) `
        ("rows with a cell but no sheet: " + $sheetless.Count)

    $checkFails = Get-XRayCheckFailures
    if ($checkFails -eq 0) {
        Complete-Test -Pass -Detail ("TxB=[{0}] CA_Scalar=[{1}] TxRetQArray=[{2}]" -f `
                                     ($txb -join ' '), ($vba -join ' '), ($ret -join ' '))
    }
    Complete-Test -Fail -Detail "$checkFails case(s) failed"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' '))
}
