$case = @{ Name='udf-volatile-recalc-twice'
     Modules=@{
       'M'=@'
Public Function VolUdf(ByVal n As Double) As Double
    Application.Volatile
    VolUdf = n + 1
End Function
Public Sub TwoCalcs()
    Application.CalculateFull
    Application.CalculateFull
End Sub
'@
     }
     FillFormula=@{ Range='A1:A20'; Formula='=VolUdf(ROW())' }
     Trigger=@{ Kind='Run'; Name='TwoCalcs' }
     Expect={ param($t)
        $c = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'VolUdf' }).Count
        if ($c -lt 40) { return "expected >=40 activations over two recalcs, got $c" }
        $null }
     Why='a volatile UDF forced through two full recalculations in one armed session' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
