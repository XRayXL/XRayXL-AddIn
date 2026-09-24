# A cell UDF that touches the object model under recalculation is not stamped `threw`, and its
# caller still resolves to its cell; its helper nests under it.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Function P6_Udf(ByVal x As Long) As Variant
    Application.Volatile
    Dim c As Range
    Set c = Application.Caller                                  ' object-model: who called me
    Dim other As Variant
    other = ThisWorkbook.Worksheets("S1").Range("Z1").Value     ' object-model: read a cell
    P6_Udf = x + P6_Helper(x) + CLng(other)
End Function

Public Function P6_Helper(ByVal x As Long) As Long
    P6_Helper = x * 10
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P6' @(
        @{ Kind=1; Name='P6Case'; Code=$moduleCode }
    ) @{
        'A1' = '=P6_Udf(5)'
    } {
        param($ws)
        $ws.Range('Z1').Value = 7
    }
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    Invoke-XRayRecalc $app
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows    = Select-BookRows (Read-TraceRows $sx.ProcId) $leaf
    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit'  -and $_.source -eq 'VBA') })
    $udf     = @($entries | Where-Object { $_.function -eq 'P6_Udf' })
    $helper  = @($entries | Where-Object { $_.function -eq 'P6_Helper' })

    Check 'udf-traced-from-cell' (($udf.Count -ge 1) -and ($udf[0].caller -eq 'cell') -and ((Get-CallerCell $udf[0]) -eq 'A1')) `
          "P6_Udf count=$($udf.Count) caller='$(if ($udf.Count){$udf[0].caller})' callerref='$(if ($udf.Count){$udf[0].callerref})'"
    $udfSpan = if ($udf.Count) { [string]$udf[0].span } else { '' }
    $helpParent = if ($helper.Count) { [string]$helper[0].parent } else { '' }
    Check 'helper-nests-under-the-udf' (($helper.Count -ge 1) -and ($helpParent -eq $udfSpan) -and ($udfSpan -ne '')) `
          "P6_Helper parent=$helpParent, P6_Udf span=$udfSpan"

    $notRet = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'object-model-touches-do-not-read-as-errors' ($notRet.Count -eq 0) `
          ("not returned: " + (@($notRet | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "object-model-touching cell UDF traced clean: P6_Udf (+helper), caller=cell, all returned"
}
catch { Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' })) }
