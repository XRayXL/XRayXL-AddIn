$case = @{ Name='udf-nested-three-deep'
     Modules=@{
       'M'=@'
Private cA As Long
Private cB As Long
Private cC As Long
Public Function NestA(ByVal x As Double) As Double
    cA = cA + 1
    NestA = NestB(x) + 1
End Function
Public Function NestB(ByVal x As Double) As Double
    cB = cB + 1
    NestB = NestC(x) + 1
End Function
Public Function NestC(ByVal x As Double) As Double
    cC = cC + 1
    NestC = x * 2
End Function
Public Function XR_Counts() As String
    XR_Counts = "NestA=" & cA & ";NestB=" & cB & ";NestC=" & cC
End Function
'@
     }
     FillFormula=@{ Range='A1:A50'; Formula='=NestA(ROW())' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'NestA','NestB','NestC'; if ($why) { return $why }
        if ($t.maxDepth -lt 3) { return "expected depth >=3, got $($t.maxDepth)" }
        # How many times Excel evaluates the sheet is Excel's business; that every call is
        # traced is ours. VBA counted them itself.
        Test-TracedCallsMatchCounters $t }
     Why='UDFs calling UDFs three levels down, from many cells at once' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
