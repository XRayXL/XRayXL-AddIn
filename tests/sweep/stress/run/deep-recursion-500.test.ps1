$case = @{ Name='deep-recursion-500'
     Modules=@{
       'M'=@'
Public Sub Deep(ByVal n As Long)
    If n > 0 Then Deep n - 1
End Sub
Public Sub Go()
    Deep 500
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted at depth" }
        if ($t.stackGrowFailures -gt 0) { return "the frame stack could not grow ($($t.stackGrowFailures) time(s))" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "frames leaked: $($t.framesOpened)/$($t.framesClosed)" }
        if ($t.maxDepth -ne 502) { return "Go and Deep 500 down to Deep 0 are 502 deep, but maxDepth is $($t.maxDepth)" }
        if ($t.recursions -ne 500) { return "expected 500 recursions, got $($t.recursions)" }
        $null }
     # Past the frame stack's first 256 frames it grows, so every activation has its rows.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(0..500 | ForEach-Object { @{ Function='Deep'; Args="a1:Long=$(500 - $_)"; Depth="$($_ + 2)"; Parent=$_; Outcome='returned' } })
     Why='a recursion deeper than the frame stack first holds: every level must still be recorded' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
