$case = @{ Name='result-float-arithmetic'
     Modules=@{
       'M'=@'
Public Function FpSum(ByVal n As Double) As Double
    Dim i As Long
    Dim s As Double
    s = 0
    For i = 1 To 100
        s = s + (1# / (i * 1#)) * n
    Next i
    FpSum = s
End Function
Public Sub Go()
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets("S1")
    ws.Range("D1").Value = FpSum(1)
    ws.Range("D2").Value = FpSum(2)
    ws.Range("D3").Value = 1.5 * 2.5
    ws.Range("D4").Value = 1# / 3#
End Sub
'@
     }
     VerifyCells=@{ 'D3'='3.75'; 'D4'='0.333333333333333' }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.statements -lt 100) { return "expected the VBA to have run, got $($t.statements) statements" }
        $null }
     Why='heavy Double arithmetic under tracing: the answers must be bit-identical to what Excel computes untraced. This is the case that would catch a thunk clobbering xmm0' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
