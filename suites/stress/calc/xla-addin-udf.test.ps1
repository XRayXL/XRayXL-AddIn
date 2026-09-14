$case = @{ Name='xla-addin-udf'
     FillFormula=@{ Range='A1:A30'; Formula='=''AddOne.xlam''!AddinUdf(ROW())' }
     Deps=@(
       @{ Name='AddOne'; IsAddin=$true; Modules=@{
         'MA'=@'
Public Function AddinUdf(ByVal n As Double) As Double
    AddinUdf = n * 13
End Function
'@
       } }
     )
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $e = $t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'AddinUdf' } | Select-Object -First 1
        if (-not $e) { return "the add-in UDF was never traced" }
        if ($e.module -notmatch 'AddOne') {
            return "add-in code attributed to the wrong file: [$($e.module)]" }
        $null }
     Why='an .xlam add-in providing a UDF that a workbook uses UNQUALIFIED' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
