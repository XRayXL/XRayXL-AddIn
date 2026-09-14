$case = @{ Name='udf-takes-range-argument'
     Modules=@{
       'M'=@'
Public Function RangeUdf(ByVal r As Range) As Double
    Dim z As Double
    z = r.Cells(1, 1).Value2
    RangeUdf = z
End Function
'@
     }
     Cells=@{ 'B1'='17'; 'A1'='=RangeUdf(B1:B1)' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'RangeUdf' } | Select-Object -First 1
        if (-not $e) { return "no row for RangeUdf" }
        if ($e.typetext -notmatch 'Object|\?') { return "unexpected signature [$($e.typetext)]" }
        $null }
     Why='a UDF taking a Range: the argument is an object reference, not a value' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
