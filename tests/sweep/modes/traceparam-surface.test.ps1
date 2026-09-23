# XRayXL_SetTraceParam / XRayXL_GetTraceParam -- the whole surface.
#
# One setting per call, addressed by (Source, Name). Omitting Source addresses
# BOTH sources, which works because an XLL argument that is not supplied
# arrives as xltypeMissing whatever its POSITION -- so the leading argument
# can be omitted without a placeholder. Get returns a scalar when both are
# given and a grid otherwise, and the four shapes are asserted here because
# their sizes are what a sheet formula spills into.
#
# The refusals matter as much as the settings: a value the tracer cannot
# honour must fail loudly and change NOTHING, never half-apply.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # 1. Set one thing, read it back. The other source is read first: what is asserted is
    # independence, that changing XLL leaves VBA alone, not what either defaults to.
    $vbaBefore = [string](Get-XRayTraceParam $sx 'VBA' 'DEPTH')

    $echo = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'TOP'
    Check 'set-xll-depth-echoes' ($echo -match 'XLL DEPTH=TOP') $echo
    $v = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    Check 'get-xll-depth-scalar' ($v -eq 'TOP') "got '$v'"

    # The OTHER source must be untouched -- one call changes one thing.
    $v = [string](Get-XRayTraceParam $sx 'VBA' 'DEPTH')
    Check 'set-one-source-leaves-the-other' ($v -eq $vbaBefore) "VBA DEPTH was '$vbaBefore' before the XLL set and is '$v' after"

    # ---- 2. OMITTED SOURCE ADDRESSES BOTH ---------------------------------
    $echo = Set-XRayTraceParam $sx $null 'ARGS' $false
    Check 'set-both-sources-echoes' ($echo -match 'ARGS: XLL=FALSE VBA=FALSE') $echo
    $x = [string](Get-XRayTraceParam $sx 'XLL' 'ARGS')
    $b = [string](Get-XRayTraceParam $sx 'VBA' 'ARGS')
    Check 'omitted-source-set-both' (($x -eq 'FALSE') -and ($b -eq 'FALSE')) "XLL=$x VBA=$b"
    [void](Set-XRayTraceParam $sx $null 'ARGS' $true)

    # 3. The four Get shapes, as a sheet sees them. Application.Run flattens an XLL array to a
    # 1-D object[], so the shape is asserted through a formula, which is also how a user asks.
    $shapeBook = Join-Path $sx.WorkDir ("ParamShape_{0}.xlsx" -f $sx.ProcId)
    Close-OwnLeftover $app (Split-Path $shapeBook -Leaf)
    Remove-Item $shapeBook -ErrorAction SilentlyContinue
    $booksRef = $app.Workbooks
    $swb = $booksRef.Add(); try { $swb.EnableAutoRecover = $false } catch {}
    $sws = $swb.Worksheets.Item(1)
    # Entered as an array formula over an oversized range: Excel fills the array it was given
    # and pads the rest with #N/A, so the populated corner proves the width and height and the
    # #N/A cells prove it is not larger. A COM-set dynamic-array spill shows only its top-left
    # cell, which proves nothing about the shape.
    $sws.Range('A1:C7').FormulaArray = '=XRayXL_GetTraceParam("XLL")'
    $sws.Range('E1:G3').FormulaArray = '=XRayXL_GetTraceParam(,"DEPTH")'
    $sws.Range('J1:M13').FormulaArray = '=XRayXL_GetTraceParam()'
    Invoke-XRayRecalc $app

    function CellText($r) { return (Get-XRayCellText $sws.Range($r)) }

    # Source only -> exactly 5 rows x 2 cols: Name, Value. The last row proves
    # the grid covers the whole settings list (5 settings, DEPTH..BREAKPOINTS).
    $ok = (CellText 'A1') -eq 'DEPTH' -and (CellText 'A4') -eq 'OBJECTS' -and (CellText 'A5') -eq 'BREAKPOINTS' `
          -and (CellText 'B1') -ne '' -and (CellText 'C1') -eq '#N/A' -and (CellText 'A6') -eq '#N/A'
    Check 'array-source-only-is-5x2' $ok ("A1='{0}' A5='{1}' B1='{2}' C1='{3}' A6='{4}'" -f `
          (CellText 'A1'), (CellText 'A5'), (CellText 'B1'), (CellText 'C1'), (CellText 'A6'))

    # Name only -> exactly 2 rows x 2 cols: Source, Value.
    $ok = (CellText 'E1') -eq 'XLL' -and (CellText 'E2') -eq 'VBA' -and (CellText 'F1') -ne '' `
          -and (CellText 'G1') -eq '#N/A' -and (CellText 'E3') -eq '#N/A'
    Check 'array-name-only-is-2x2' $ok ("E1='{0}' E2='{1}' F1='{2}' G1='{3}' E3='{4}'" -f `
          (CellText 'E1'), (CellText 'E2'), (CellText 'F1'), (CellText 'G1'), (CellText 'E3'))

    # Neither -> exactly 10 rows x 3 cols: Source, Name, Value (2 sources x 5 settings).
    $ok = (CellText 'J1') -eq 'XLL' -and (CellText 'J5') -eq 'XLL' -and (CellText 'J6') -eq 'VBA' -and (CellText 'L1') -ne '' `
          -and (CellText 'M1') -eq '#N/A' -and (CellText 'J11') -eq '#N/A'
    Check 'array-everything-is-10x3' $ok ("J1='{0}' J6='{1}' L1='{2}' M1='{3}' J11='{4}'" -f `
          (CellText 'J1'), (CellText 'J6'), (CellText 'L1'), (CellText 'M1'), (CellText 'J11'))

    # The CONTENT still arrives whole through Application.Run, in row-major
    # order -- only the shape is lost, so the values remain assertable there.
    $flat = @(ConvertTo-XRayGrid (Get-XRayTraceParam $sx 'XLL' $null))[0]
    Check 'run-returns-all-cells' ($flat.Count -eq 10) "$($flat.Count) cells"
    Check 'run-cells-row-major' (($flat[0] -eq 'DEPTH') -and ($flat[2] -eq 'ARGS') -and
                                 ($flat[4] -eq 'RETVAL') -and ($flat[6] -eq 'OBJECTS') -and ($flat[8] -eq 'BREAKPOINTS')) ($flat -join ',')

    # ---- 4. REFUSALS: loud, and nothing changes ---------------------------
    $before = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')

    $e = Set-XRayTraceParam $sx 'NOPE' 'DEPTH' 'ALL'
    Check 'bad-source-refused' ($e -match '#Err') $e
    $e = Set-XRayTraceParam $sx 'XLL' 'NOSUCH' 'ALL'
    Check 'bad-name-refused' ($e -match '#Err') $e
    $e = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'SIDEWAYS'
    Check 'bad-depth-refused' ($e -match '#Err') $e
    $e = Set-XRayTraceParam $sx 'XLL' 'ARGS' 'PERHAPS'
    Check 'bad-bool-refused' ($e -match '#Err') $e

    # With no Source the recording-wide settings are names too, so the refusal lists all nine.
    $e = Set-XRayTraceParam $sx $null 'NOSUCH' 'ALL'
    Check 'no-source-refusal-names-all-nine' (($e -match 'BUFFERSIZE') -and ($e -match 'BUFFERWHENFULL') -and ($e -match 'FORMAT') -and ($e -match 'LOGLEVEL') -and ($e -match 'OBJECTS') -and ($e -match 'BREAKPOINTS')) $e

    # A getter refusal arrives whole, not cut at the grid's cell width.
    $e = [string](Get-XRayTraceParam $sx 'NOPE' 'DEPTH')
    Check 'getter-refusal-is-whole' ($e -eq '#Err - Source must be XLL or VBA, or omit for both') $e
    $e = [string](Get-XRayTraceParam $sx 'XLL' 'NOSUCH')
    Check 'getter-name-refusal-is-whole' ($e -eq '#Err - Name must be DEPTH, ARGS, RETVAL, OBJECTS or BREAKPOINTS') $e

    # DROP leaves one hole per dropped row, and the echo says so.
    $full = [string](Get-XRayTraceParam $sx $null 'BUFFERWHENFULL')
    $e = Set-XRayTraceParam $sx $null 'BUFFERWHENFULL' 'DROP'
    Check 'drop-echo-says-each-row-leaves-a-hole' ($e -match 'each leaving a hole in the input column') $e
    [void](Set-XRayTraceParam $sx $null 'BUFFERWHENFULL' $full)

    $after = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    Check 'refusals-changed-nothing' ($after -eq $before) "was '$before', now '$after'"

    # CALLER is not a setting: the calling cell is always resolved, since the VBA tracer needs
    # it to place an error that escapes into a cell. The name is refused like any other unknown
    # one.
    $e = Set-XRayTraceParam $sx 'XLL' 'CALLER' $false
    Check 'caller-is-not-a-setting' ($e -match '#Err') $e

    # ---- 5. REFUSES WHILE ARMED -------------------------------------------
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    if (Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark) {
        $e = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF'
        Check 'refused-while-armed' ($e -match '#Err - cannot change settings while armed') $e
        # The QUERY is not a change and must still answer.
        $v = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
        Check 'get-answers-while-armed' ($v -eq $before) "got '$v'"
    }
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    # ---- 6. REFUSES FROM A CELL -------------------------------------------
    # A formula that reconfigures the tracer on every recalc is a foot-gun,
    # and xlfCaller says exactly who is asking.
    $bookPath = Join-Path $sx.WorkDir ("ParamCell_{0}.xlsx" -f $sx.ProcId)
    Close-OwnLeftover $app (Split-Path $bookPath -Leaf)
    Remove-Item $bookPath -ErrorAction SilentlyContinue
    $booksRef = $app.Workbooks
    $wb = $booksRef.Add(); try { $wb.EnableAutoRecover = $false } catch {}
    $ws = $wb.Worksheets.Item(1)
    $ws.Range('A1').Formula = '=XRayXL_SetTraceParam("XLL","DEPTH","OFF")'
    Invoke-XRayRecalc $app
    $cell = Get-XRayCellText $ws.Range('A1')
    Check 'refused-from-a-cell' ($cell -match '#Err') "cell says '$cell'"
    $v = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    Check 'cell-call-changed-nothing' ($v -eq $before) "DEPTH is '$v', expected '$before'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'set, get in four shapes, refusals loud and inert, armed and cell guards'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
