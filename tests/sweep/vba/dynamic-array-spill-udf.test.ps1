# A UDF whose array result spills names its anchor cell as caller and decodes its array return,
# which scalar UDF tests never reach; a helper per element makes the nesting real.
# Spill needs a recent Excel; the UDF's own return is the same whether or not it spills.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Function P8_Spill(ByVal n As Long) As Variant
    Dim a() As Long
    ReDim a(1 To n)
    Dim i As Long
    For i = 1 To n
        a(i) = P8_Weight(i)
    Next i
    P8_Spill = a
End Function

Public Function P8_Weight(ByVal i As Long) As Long
    P8_Weight = i * 3
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P8' @(
        @{ Kind=1; Name='P8Case'; Code=$moduleCode }
    ) @{
        'A1' = '=P8_Spill(3)'
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
    $spill   = @($entries | Where-Object { $_.function -eq 'P8_Spill' })
    $weight  = @($entries | Where-Object { $_.function -eq 'P8_Weight' })

    Check 'spill-udf-traced-from-cell' (($spill.Count -ge 1) -and ($spill[0].caller -eq 'cell') -and ((Get-CallerCell $spill[0]) -eq 'A1')) `
          "P8_Spill count=$($spill.Count) caller='$(if($spill.Count){$spill[0].caller})' callerref='$(if($spill.Count){$spill[0].callerref})'"
    Check 'helper-called-once-per-element' ($weight.Count -eq 3) "P8_Weight entries: $($weight.Count) (expected 3)"

    $notRet = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'every-frame-returned' ($notRet.Count -eq 0) `
          ("not returned: " + (@($notRet | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    # a Long(1 To 3) in a Variant renders by its element type, as R_ArrVar does
    $spillExit = @($exits | Where-Object { $_.function -eq 'P8_Spill' })
    $ret = if ($spillExit.Count) { [string]$spillExit[0].ret } else { '(no row)' }
    Check 'array-return-rendered' ($ret -eq 'Long[1..3]{3,6,9}') "P8_Spill ret='$ret', expected Long[1..3]{3,6,9}"

    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed (ret='$ret')" }
    Complete-Test -Pass -Detail "spill UDF traced: P8_Spill (+3 helpers) from A1, ret='$ret'"
}
catch { Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' })) }
