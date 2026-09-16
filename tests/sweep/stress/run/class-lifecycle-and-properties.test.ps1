$case = @{ Name='class-lifecycle-and-properties'
     Modules=@{
       'M'=@'
Public Sub Go()
    Dim c As CThing
    Dim i As Long
    For i = 1 To 20
        Set c = New CThing
        c.Val = i
        c.Work
        Set c = Nothing
    Next i
End Sub
'@
     }
     Classes=@{
       'CThing'=@'
Private mV As Long
Private Sub Class_Initialize()
    mV = 1
End Sub
Private Sub Class_Terminate()
    mV = 0
End Sub
Public Property Let Val(ByVal v As Long)
    mV = v
End Property
Public Property Get Val() As Long
    Val = mV
End Property
Public Sub Work()
    mV = mV + 1
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'Class_Initialize','Work'; if ($why) { return $why }
        $null }
     # Twenty rounds of New (Class_Initialize), Property Let Val, Work, and Nothing (Class_Terminate),
     # each inside Go. Property Get is never called.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(1..20 | ForEach-Object {
               @{ Function='Class_Initialize'; Depth='2'; Parent=0; Outcome='returned' }
               @{ Function='Val'; Args="a1:Long=$_"; Depth='2'; Parent=0; Outcome='returned' }
               @{ Function='Work'; Depth='2'; Parent=0; Outcome='returned' }
               @{ Function='Class_Terminate'; Depth='2'; Parent=0; Outcome='returned' } })
     Why='a class module: Initialize, Property Let/Get, a method, and Terminate' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
