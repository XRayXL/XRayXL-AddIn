# A VBA UDF DEFINED IN ONE WORKBOOK, CALLED FROM ANOTHER.
#
# Under an armed CalculateFull the engine recalculates EVERY open workbook -- the
# realism Select-BookRows exists for. A cross-book UDF call is the case where the
# function lives in book A but the CALLER is a cell in book B, so identity
# resolution and caller attribution are pulled in two directions at once. The
# function must still be named (from book A) and the caller must resolve to book
# B's cell.
#
# Cross-book UDF references are version- and setting-sensitive; if the call does
# not evaluate on this Excel the test SKIPS rather than failing for a reason that
# is not the tracer.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCodeA = @'
Public Function P7_Ext(ByVal n As Long) As Long
    P7_Ext = n * 7
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # Book A provides the UDF and stays open, so book B's cross-book formula
    # resolves when it is written and again when B is reopened.
    New-XRayMacroBook $sx 'P7A' @(@{ Kind = 1; Name = 'P7ACase'; Code = $moduleCodeA })
    $leafA = (Get-XRayMacroBook 'P7A').Leaf
    New-XRayMacroBook $sx 'P7B' -Prepare {
        param($sheet)
        try { $sheet.Range('A1').Formula = "=$leafA!P7_Ext(3)" }
        catch { Complete-Test -Skip -Detail "this Excel would not accept a cross-book UDF formula: $($_.Exception.Message)" }
    }
    $bookB = Get-XRayMacroBook 'P7B'
    $leafB = $bookB.Leaf; $wsB = $bookB.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    Invoke-XRayRecalc $app
    $cellNow = Get-XRayCellText $wsB.Range('A1')
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    if ($cellNow -like '#*') { Complete-Test -Skip -Detail "cross-book UDF did not evaluate on this Excel (A1='$cellNow'); not a tracer question" }

    $rows  = @(Read-TraceRows $sx.ProcId)
    $ext   = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'P7_Ext') })

    Check 'cross-book-udf-computed' ($cellNow -eq '21') "A1='$cellNow', expected 21"
    Check 'cross-book-udf-was-traced' ($ext.Count -ge 1) "P7_Ext entries: $($ext.Count)"
    if ($ext.Count) {
        Check 'caller-is-the-calling-cell' ($ext[0].caller -eq 'cell') "caller='$($ext[0].caller)'"
        # The caller/cell belong to book B (the CALLER), whatever book defined the UDF.
        Check 'caller-cell-belongs-to-the-calling-book' ((Get-CallerSheet $ext[0]) -like "*$leafB*") `
              "callerref='$($ext[0].callerref)' (should name the calling book $leafB)"
    }
    Write-Output ("P7_Ext rows=$($ext.Count) caller='$(if($ext.Count){$ext[0].caller})' callerref='$(if($ext.Count){$ext[0].callerref})'")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "cross-book UDF traced: P7_Ext from $leafA, called by $leafB!A1"
}
catch { Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' })) }
