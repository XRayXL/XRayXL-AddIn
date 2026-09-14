# XRayXL_GetTraceSummary -- what has actually been traced.
#
# One row per function that was CALLED, with its call count, from both
# sources. The counts are not gathered by this function: both sides already
# keep them on the hot path (Target::calls, Proc::calls), so it only formats
# what is already true -- building an array on the hot path would break the
# no-allocation rule.
#
# Read LIVE, while armed, from a worksheet. That is safe because entries are
# only ever added and counters only rise, so a racing reader sees a slightly
# stale count and never garbage.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    function Summary($filter) {
        $v = if ($null -eq $filter) { $app.Run('XRayXL_GetTraceSummary') }
             else { $app.Run('XRayXL_GetTraceSummary', $filter) }
        return (ConvertFrom-XRayTraceSummary $v)
    }

    $srcM = @'
Public Function S_Add(ByVal a As Double, ByVal b As Double) As Double
    S_Add = a + b
End Function
Public Function S_Other() As Double
    S_Other = 1
End Function
'@
    # FIVE calls to S_Add, one to S_Other, one to the XLL's TxB -- distinct
    # counts, so a summary that mixed them up could not pass by accident.
    New-XRayMacroBook $sx 'Summary' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'B1' = '=S_Other()'
        'C1' = '=TxB(2,3)'
    } {
        param($ws)
        1..5 | ForEach-Object { $ws.Range("A$_").Formula = '=S_Add(1,2)' }
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    [void](Wait-LogLine $paths.Log 'VBA tracing: ' $mark)
    Invoke-XRayRecalc $app

    # ---- READ IT WHILE STILL ARMED ----------------------------------------
    $sum = Summary $null
    $h = $sum.Header
    Check 'has-header' (($null -ne $h) -and ($h.Source -eq 'Source') -and ($h.Function -eq 'Function') -and ($h.Calls -eq 'Calls')) `
          $(if ($h) { "$($h.Source)|$($h.Module)|$($h.Function)|$($h.Calls)" } else { 'no header row' })

    $data  = $sum.Rows
    $add   = @($data | Where-Object { $_.Function -eq 'S_Add' })   | Select-Object -First 1
    $other = @($data | Where-Object { $_.Function -eq 'S_Other' }) | Select-Object -First 1
    $txb   = @($data | Where-Object { $_.Function -eq 'TxB' })     | Select-Object -First 1
    $seen  = @($data | ForEach-Object { $_.Function }) -join ','

    Check 'vba-function-listed' ($null -ne $add) "functions: $seen"
    if ($add) {
        Check 'vba-source-is-VBA' ($add.Source -eq 'VBA') "source='$($add.Source)'"
        # Five cells call it; the sheet may recalculate again while this runs,
        # and a count that only rises guarantees at least five, not exactly five.
        Check 'vba-count-at-least-five' ([double]$add.Calls -ge 5) "calls=$($add.Calls)"
        Check 'vba-module-named' ($add.Module -match 'M') "module='$($add.Module)'"
    }
    if ($add -and $other) {
        Check 'counts-are-per-function' ([double]$other.Calls -lt [double]$add.Calls) `
              "S_Other=$($other.Calls) vs S_Add=$($add.Calls)"
    }
    Check 'xll-function-listed' ($null -ne $txb) "functions: $seen"
    if ($txb) { Check 'xll-source-is-XLL' ($txb.Source -eq 'XLL') "source='$($txb.Source)'" }

    # ---- THE WILDCARD FILTER ----------------------------------------------
    $names = @((Summary 'S_*').Rows | ForEach-Object { $_.Function })
    Check 'filter-keeps-matches' (($names -contains 'S_Add') -and ($names -contains 'S_Other')) ($names -join ',')
    Check 'filter-drops-non-matches' ($names -notcontains 'TxB') ($names -join ',')

    $names = @((Summary 'S_A*').Rows | ForEach-Object { $_.Function })
    Check 'filter-is-a-wildcard-not-a-prefix' (($names -contains 'S_Add') -and ($names -notcontains 'S_Other')) ($names -join ',')

    # A filter that matches nothing SAYS so, rather than returning an empty
    # grid that reads as a broken formula.
    $names = @((Summary 'ZZZ_NOTHING*').Rows | ForEach-Object { $_.Function })
    Check 'empty-result-says-so' (($names -join ',') -match 'nothing traced') ($names -join ',')

    # ---- SHAPE, THROUGH A SHEET -------------------------------------------
    # Application.Run flattens, so the column count is asserted where it is
    # real: an oversized array formula pads with #N/A.
    $ws.Range('F1:K3').FormulaArray = '=XRayXL_GetTraceSummary("S_A*")'
    Invoke-XRayRecalc $app
    $h = @('F1','G1','H1','I1') | ForEach-Object { Get-XRayCellText $ws.Range($_) }
    Check 'sheet-header-4-columns' (($h[0] -eq 'Source') -and ($h[3] -eq 'Calls')) ($h -join '|')
    $j1 = Get-XRayCellText $ws.Range('J1'); $h2 = Get-XRayCellText $ws.Range('H2')
    Check 'sheet-is-exactly-4-wide' ($j1 -eq '#N/A') "J1='$j1'"
    Check 'sheet-lists-the-match' ($h2 -eq 'S_Add') "H2='$h2'"

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'both sources listed with live counts, wildcard filter, empty case loud, 4 columns'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
