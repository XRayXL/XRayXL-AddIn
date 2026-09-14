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
     Why='a class module: Initialize, Property Let/Get, a method, and Terminate' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
