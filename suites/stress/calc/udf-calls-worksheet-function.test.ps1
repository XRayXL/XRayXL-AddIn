$case = @{ Name='udf-calls-worksheet-function'
     Modules=@{
       'M'=@'
Public Function WsfUdf(ByVal n As Double) As Double
    WsfUdf = Application.WorksheetFunction.Max(n, 10)
End Function
'@
     }
     FillFormula=@{ Range='A1:A30'; Formula='=WsfUdf(ROW())' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'WsfUdf'; if ($why) { return $why }
        $null }
     Why='a UDF calling back into Excel via WorksheetFunction while being traced' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
