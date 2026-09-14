$case = @{ Name='empty-and-oneline-procedures'
     Modules=@{
       'M'=@'
Public Sub E1()
End Sub
Public Sub E2()
End Sub
Public Sub E3()
End Sub
Public Sub One()
    Dim z As Long
End Sub
Public Sub Go()
    E1
    E2
    E3
    One
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 5) { return "expected 5 procedures, got $($t.procedures)" }
        $null }
     Why='procedures with no body at all, mixed with single-statement ones' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
