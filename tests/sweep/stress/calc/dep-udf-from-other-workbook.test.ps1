$case = @{ Name='dep-udf-from-other-workbook'
     Cells=@{ 'A1'='=''DepB.xlsm''!OtherUdf(6)' }
     Deps=@(
       @{ Name='DepB'; Modules=@{
         'MB'=@'
Public Function OtherUdf(ByVal n As Double) As Double
    OtherUdf = n * 7
End Function
'@
       } }
     )
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'OtherUdf' } | Select-Object -First 1
        if (-not $e) { return "the other workbook''s UDF was never traced" }
        if ($e.module -notmatch 'DepB') {
            return "UDF attributed to the wrong workbook: [$($e.module)]" }
        $null }
     Why='a formula in book A calling a UDF defined in book B: the row must name B, not A' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
