$case = @{ Name='err-goto-cleanup-pattern'
     Modules=@{
       'M'=@'
Public Sub Work(ByVal n As Long)
    On Error GoTo Cleanup
    If n Mod 3 = 0 Then Err.Raise 5201
    Exit Sub
Cleanup:
    Resume Next
End Sub
Public Sub Go()
    Dim i As Long
    For i = 1 To 200
        Work i
    Next i
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 2) { return "expected 2 procedures, got $($t.procedures)" }
        $null }
     # Work raises and handles inside its own frame, so nothing crosses a boundary: all returned.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(1..200 | ForEach-Object { @{ Function='Work'; Args="a1:Long=$_"; Depth='2'; Parent=0; Outcome='returned' } })
     Why='the ordinary On Error GoTo Cleanup / Resume shape, run many times' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
