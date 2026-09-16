$case = @{ Name='udf-returns-array'
     Modules=@{
       'M'=@'
Private cM As Long
Public Function MakeArr() As Variant
    Dim a(1 To 500) As Double
    Dim i As Long
    cM = cM + 1
    For i = 1 To 500
        a(i) = i * 1.5
    Next i
    MakeArr = a
End Function
Public Function XR_Counts() As String
    XR_Counts = "MakeArr=" & cM
End Function
'@
     }
     Cells=@{ 'A1'='=SUM(MakeArr())' }
     Trigger=@{ Kind='Calc' }
     Counters='XR_Counts'
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'MakeArr'; if ($why) { return $why }
        Test-TracedCallsMatchCounters $t }
     Why='a UDF returning a 500-element array into a spilled formula' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
