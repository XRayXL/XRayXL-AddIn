$case = @{ Name='udf-optional-and-paramarray'
     Modules=@{
       'M'=@'
Private cF As Long
Public Function Flex(ByVal a As Double, ParamArray rest() As Variant) As Double
    Dim z As Double
    cF = cF + 1
    z = a
    If UBound(rest) >= 0 Then z = z + 1
    Flex = z
End Function
Public Function XR_Counts() As String
    XR_Counts = "Flex=" & cF
End Function
'@
     }
     Cells=@{ 'A1'='=Flex(1)'; 'A2'='=Flex(1,2)'; 'A3'='=Flex(1,2,3,4)' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Flex'; if ($why) { return $why }
        Test-TracedCallsMatchCounters $t }
     Why='a UDF with a ParamArray, called from cells with varying arity. NOTE: VBA forbids Optional and ParamArray in one signature -- trying it is a COMPILE error that wedges Excel, which is how the per-case deadline earned its keep' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
