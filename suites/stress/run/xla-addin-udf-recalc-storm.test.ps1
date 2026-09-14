$case = @{ Name='xla-addin-udf-recalc-storm'
     Modules=@{
       'M'=@'
Public Sub TwoCalcs()
    Application.CalculateFull
    Application.CalculateFull
End Sub
'@
     }
     FillFormula=@{ Range='A1:A150'; Formula='=''AddFour.xlam''!StormUdf(ROW())' }
     Deps=@(
       @{ Name='AddFour'; IsAddin=$true; Modules=@{
         'MA'=@'
Public Function StormUdf(ByVal n As Double) As Double
    Application.Volatile
    StormUdf = n + 0.5
End Function
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='TwoCalcs' }
     Expect={ param($t)
        $c = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'StormUdf' }).Count
        if ($c -lt 150) { return "expected >=150 activations, got $c" }
        if ($t.unnamed -gt 0) { return "$($t.unnamed) add-in procedure(s) lost their name under churn" }
        $null }
     Why='an add-in UDF across 150 cells recalculated twice: identity under sustained churn' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
