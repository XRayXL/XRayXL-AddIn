$case = @{ Name='udf-calls-worksheet-function'
     Modules=@{
       'M'=@'
Private cW As Long
Public Function WsfUdf(ByVal n As Double) As Double
    cW = cW + 1
    WsfUdf = Application.WorksheetFunction.Max(n, 10)
End Function
Public Function XR_Counts() As String
    XR_Counts = "WsfUdf=" & cW
End Function
'@
     }
     FillFormula=@{ Range='A1:A30'; Formula='=WsfUdf(ROW())' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'WsfUdf'; if ($why) { return $why }
        # WorksheetFunction.Max is Excel's, not a VBA procedure: only WsfUdf is counted.
        Test-TracedCallsMatchCounters $t }
     Why='a UDF calling back into Excel via WorksheetFunction while being traced' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
