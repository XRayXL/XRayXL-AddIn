$case = @{ Name='button-raises-handled'
     Modules=@{
       'M'=@'
Public Sub BtnRaise()
    On Error GoTo H
    Inner
    Exit Sub
H:
    Dim z As Long
    z = Err.Number
End Sub
Public Sub Inner()
    Err.Raise 5501
End Sub
'@
     }
     Shapes=@(@{ Index=0; Name='B1'; OnAction='BtnRaise' })
     Trigger=@{ Kind='Button'; Name='B1' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'BtnRaise'; if ($why) { return $why }
        $null }
     Calls=@(
        @{ Function='BtnRaise'; Depth='1'; Parent=-1; Outcome='handled' }
        @{ Function='Inner';    Depth='2'; Parent=0;  Outcome='threw' } )
     Why='a button macro that raises and handles: unwinding on the macro-dispatch path' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
