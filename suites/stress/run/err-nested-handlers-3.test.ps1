$case = @{ Name='err-nested-handlers-3'
     Modules=@{
       'M'=@'
Public Sub Inner()
    Err.Raise 5101
End Sub
Public Sub Mid1()
    On Error GoTo H
    Inner
    Exit Sub
H:
    Err.Raise 5102
End Sub
Public Sub Outer()
    On Error GoTo H
    Mid1
    Exit Sub
H:
    Err.Raise 5103
End Sub
Public Sub Go()
    On Error Resume Next
    Outer
    Err.Clear
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 4) { return "expected 4 procedures, got $($t.procedures)" }
        $null }
     Why='three nested handlers each catching and re-raising: unwinding through active handlers' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
