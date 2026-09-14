$case = @{ Name='err-resume-next-1000'
     Modules=@{
       'M'=@'
Public Sub Boom()
    Err.Raise 5002
End Sub
Public Sub Go()
    Dim i As Long
    On Error Resume Next
    For i = 1 To 1000
        Boom
        Err.Clear
    Next i
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.statements -lt 2000) { return "expected >=2000 statements, got $($t.statements)" }
        $null }
     Why='a thousand raised-and-swallowed errors in a loop: unwinding a thousand times' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
