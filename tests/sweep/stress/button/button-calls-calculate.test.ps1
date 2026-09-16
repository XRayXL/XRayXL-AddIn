$case = @{ Name='button-calls-calculate'
     Modules=@{
       'M'=@'
Public cBtn As Long
Public cUdf As Long
Public Sub BtnCalc()
    cBtn = cBtn + 1
    Application.CalculateFull
End Sub
Public Function CalcUdf(ByVal n As Double) As Double
    cUdf = cUdf + 1
    CalcUdf = n * 3
End Function
Public Function XR_Counts() As String
    XR_Counts = "BtnCalc=" & cBtn & ";CalcUdf=" & cUdf
End Function
'@
     }
     Cells=@{ 'A1'='=CalcUdf(4)'; 'A2'='=CalcUdf(5)' }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='BtnCalc' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'BtnCalc','CalcUdf'; if ($why) { return $why }
        # How many cells CalculateFull evaluates is Excel's to decide; VBA counts the calls.
        Test-TracedCallsMatchCounters $t }
     Why='a button macro that forces a full recalculation of UDFs' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
