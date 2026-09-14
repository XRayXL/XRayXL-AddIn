# WHEN Application.Caller IS A RANGE, THE TRACE MUST SAY SO.
#
# A legacy CSE array formula entered across B2:D4 is ONE formula occupying nine
# cells, and xlfCaller answers with the whole range -- xltypeSRef carrying
# rwFirst/rwLast and colFirst/colLast, or xltypeRef for the multi-area case.
#
# WHY THIS EXISTS. The decoder read rwFirst/colFirst and threw the other two
# away, so a nine-cell caller was written as "B2": not wrong enough to notice,
# and wrong in the direction that reads as a fact. Nothing caught it. The suites
# do use FormulaArray -- in modes\traceparam-surface and modes\tracesummary --
# but only on XRayXL's OWN functions, which the tracer refuses to hook
# (xllarm.cpp counts them as ownModule), so no traced call had ever had a
# multi-cell caller. 6.6 million rows of soak did not have one either: every
# array in that workload was a DYNAMIC array, which spills from a single anchor
# cell and whose caller is therefore that one cell.
#
# THE CONTROL CASE IS HALF THE TEST. A range-rendering bug that rendered
# everything as a range would pass a range-only test, so the single-cell caller
# is asserted in the same run, from the same decoder.
#
# BOTH SOURCES, because caller.h's "ONE DECODER FOR BOTH SIDES" is a claim this
# can check: the XLL hook and the VBA interpreter hook reach the same
# DecodeCaller, so a difference between them here means they do not.
#
# AND A FUNCTION THAT RETURNS AN ARRAY, entered over a range that matches it.
# TxRetQArray hands back a 2x2 xltypeMulti, so that row pair puts a multi-cell
# CALLER and a multi-cell RETURN on the same span: the caller decoder and the
# return decoder run over the same call, and a fix to one that disturbed the
# other would show here and nowhere else. The two are independent questions --
# where the formula LIVES and what it GAVE BACK -- and the trace answers them
# in different columns, which is the thing worth pinning.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
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
        # CSE, not a dynamic array: a dynamic-array formula lives in its anchor
        # cell and spills, so its caller is that ONE cell and it cannot exercise
        # this at all. FormulaArray over a range is the only thing that produces
        # a multi-cell caller.
        $ws.Range('B2:D4').FormulaArray = '=TxB(2,3)'          # XLL scalar, 9 cells
        $ws.Range('F2:F5').FormulaArray = '=CA_Scalar()'       # VBA scalar, 4 cells
        # A 2x2 return over a 2x2 range: multi-cell caller AND multi-cell return.
        $ws.Range('H2:I3').FormulaArray = '=TxRetQArray(7)'    # XLL array,  4 cells
        # The same function in ONE cell. Its caller is that cell even though what
        # it returns is still an array -- what a function GIVES BACK says nothing
        # about where it was CALLED FROM, and conflating the two is the easy
        # mistake here.
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

    # ---- a function that RETURNS an array ----------------------------------
    $ret = CellsFor 'TxRetQArray'
    Check 'array-returning-xll-array-caller-names-the-range' ([bool]($ret -contains 'H2:I3')) `
        ("TxRetQArray caller cells: " + ($ret -join ' '))
    Check 'array-returning-xll-single-cell-caller-stays-one-cell' ([bool]($ret -contains 'A4')) `
        ("TxRetQArray caller cells: " + ($ret -join ' '))

    # The return is still decoded as an array in BOTH placements -- the caller's
    # shape must not have changed what came back.
    $retRows = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.function -ieq 'TxRetQArray' })
    $retTypes = @($retRows | ForEach-Object { $_.rettype } | Sort-Object -Unique)
    Check 'array-return-still-decoded-whatever-the-caller-shape' `
        ($retRows.Count -ge 2 -and ($retTypes | Where-Object { $_ }).Count -ge 1) `
        ("{0} exit row(s), rettype(s): {1}" -f $retRows.Count, (($retTypes | Where-Object { $_ }) -join ' '))

    # NO CALLER MAY BE A BARE FIRST CELL OF AN ARRAY RANGE. This is the actual
    # defect stated directly: "B2", "F2" and "H2" are what truncation produces,
    # and none is a cell any formula in this book occupies on its own.
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
