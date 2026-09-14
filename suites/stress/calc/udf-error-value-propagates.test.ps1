$case = @{ Name='udf-error-value-propagates'
     Modules=@{
       'M'=@'
Public Function ErrUdf(ByVal n As Double) As Variant
    If n > 3 Then
        ErrUdf = CVErr(2042)
    Else
        ErrUdf = n
    End If
End Function
'@
     }
     FillFormula=@{ Range='A1:A10'; Formula='=ErrUdf(ROW())' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'ErrUdf'; if ($why) { return $why }
        $null }
     Why='a UDF returning CVErr, and another consuming the error value' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
