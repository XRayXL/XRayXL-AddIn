$case = @{ Name='err-raise-handled-deep'
     Modules=@{
       'M'=@'
Public Sub L5()
    Err.Raise 5001, "L5", "boom"
End Sub
Public Sub L4()
    L5
End Sub
Public Sub L3()
    L4
End Sub
Public Sub L2()
    L3
End Sub
Public Sub L1()
    L2
End Sub
Public Sub Go()
    On Error Resume Next
    L1
    Err.Clear
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 6) { return "expected 6 procedures, got $($t.procedures)" }
        if ($t.maxDepth -lt 6)   { return "expected depth >=6, got $($t.maxDepth)" }
        $null }
     # L5 raises through L4..L1, none handling it, to Go's Resume Next.
     Calls=@(
        @{ Function='Go'; Depth='1'; Parent=-1; Outcome='handled' }
        @{ Function='L1'; Depth='2'; Parent=0; Outcome='unwound' }
        @{ Function='L2'; Depth='3'; Parent=1; Outcome='unwound' }
        @{ Function='L3'; Depth='4'; Parent=2; Outcome='unwound' }
        @{ Function='L4'; Depth='5'; Parent=3; Outcome='unwound' }
        @{ Function='L5'; Depth='6'; Parent=4; Outcome='threw' } )
     Why='an error raised five frames down and handled at the top: frames must still balance' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
