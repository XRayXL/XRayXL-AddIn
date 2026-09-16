$case = @{ Name='result-mixed-fp-and-strings'
     Modules=@{
       'M'=@'
Public Function Mixed(ByVal n As Double, ByVal s As String) As Double
    Dim z As Double
    z = n * 0.125
    If Len(s) > 0 Then z = z + Len(s)
    Mixed = z
End Function
Public Sub Go()
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets("S1")
    ws.Range("D1").Value = Mixed(8, "abcd")
    ws.Range("D2").Value = Mixed(64, "")
    ws.Range("D3").Value = Mixed(1024, "xy")
End Sub
'@
     }
     VerifyCells=@{ 'D1'='5'; 'D2'='8'; 'D3'='130' }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.framesOpened -lt 3) { return "expected >=3 frames, got $($t.framesOpened)" }
        $null }
     # n*0.125 plus the string's length: 1+4, 8+0, 128+2.
     Calls=@(
        @{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='Mixed'; Args='a1:Double=8 a2:String="abcd"';  Ret='5';   RetType='Double'; Depth='2'; Parent=0; Outcome='returned' }
        @{ Function='Mixed'; Args='a1:Double=64 a2:String=""';     Ret='8';   RetType='Double'; Depth='2'; Parent=0; Outcome='returned' }
        @{ Function='Mixed'; Args='a1:Double=1024 a2:String="xy"'; Ret='130'; RetType='Double'; Depth='2'; Parent=0; Outcome='returned' } )
     Why='floating point interleaved with string work, which is where the CRT formatting path touches xmm hardest' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
