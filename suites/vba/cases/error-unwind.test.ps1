$case = @{ Name='error-unwind'
     Setup=@'
Public Sub T_Raise()
    Err.Raise 5
End Sub
Public Sub T_Catch()
    On Error Resume Next
    Call T_Raise
    On Error GoTo 0
End Sub
'@
     Invoke=@{ Name='T_Catch'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0)          { return "$($t.faults) faults during an error unwind" }
        # an exit with no open frame, or a frame never closed, is the shadow stack drifting
        if ($t.unmatchedExits -ne 0) { return "$($t.unmatchedExits) unmatched exit(s) -- shadow stack drifted" }
        if ($t.framesOpened -ne $t.framesClosed) { return "opened $($t.framesOpened) frame(s), closed $($t.framesClosed)" }
        $null }
     Why='THE known-weak path: an error unwind must close every frame it opened and leave no exit unmatched' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
