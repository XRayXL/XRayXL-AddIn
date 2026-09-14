$case = @{ Name='err-handled-one-up'
     Setup=@'
Public Sub T_H1C()
    Err.Raise 5, "T_H1C", "deliberate"
End Sub
Public Sub T_H1B()
    On Error GoTo done
    Call T_H1C
done:
End Sub
Public Sub T_H1A()
    Call T_H1B
    Call T_H1B
End Sub
'@
     Invoke=@{ Name='T_H1A'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) faults" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        if ($t.maxDepth -lt 3) { return "expected depth 3, got $($t.maxDepth)" }
        $null }
     Why='error unwinds ONE level and is handled -- and it happens twice, so a
          frame leaked the first time would still be open the second' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
