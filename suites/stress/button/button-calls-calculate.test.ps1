$case = @{ Name='button-calls-calculate'
     Modules=@{
       'M'=@'
Public Sub BtnCalc()
    Application.CalculateFull
End Sub
Public Function CalcUdf(ByVal n As Double) As Double
    CalcUdf = n * 3
End Function
'@
     }
     Cells=@{ 'A1'='=CalcUdf(4)'; 'A2'='=CalcUdf(5)' }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='BtnCalc' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'BtnCalc','CalcUdf'; if ($why) { return $why }
        $null }
     Why='a button macro that forces a full recalculation of UDFs' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
