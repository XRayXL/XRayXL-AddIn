$case = @{ Name='dialog-then-clean-run'
     Modules=@{
       'M'=@'
Public Sub Bad()
    Err.Raise 5903
End Sub
Public Sub Clean()
    Dim z As Long
    z = 1
    Clean2
End Sub
Public Sub Clean2()
    Dim q As Long
    q = 2
End Sub
Public Sub Go()
    On Error Resume Next
    Bad
    Err.Clear
    Clean
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.maxDepth -gt 6) { return "depth $($t.maxDepth) -- frames drifted after the error" }
        $why = Assert-VbaTraced $t 'Bad','Clean','Clean2'; if ($why) { return $why }
        $null }
     Why='a swallowed error followed by an ordinary call in the SAME armed session: the later call must not inherit the abandoned frames' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
