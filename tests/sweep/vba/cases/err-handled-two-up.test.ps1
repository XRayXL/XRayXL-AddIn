$case = @{ Name='err-handled-two-up'
     Setup=@'
Public Sub T_H2D()
    Err.Raise 5, "T_H2D", "deliberate"
End Sub
Public Sub T_H2C()
    Call T_H2D
End Sub
Public Sub T_H2B()
    Call T_H2C
End Sub
Public Sub T_H2A()
    On Error GoTo done
    Call T_H2B
done:
    Call T_Quiet
End Sub
Public Sub T_Quiet()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Name='T_H2A'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) faults" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        if ($t.maxDepth -lt 4) { return "expected depth 4 before the unwind, got $($t.maxDepth)" }
        $null }
     # T_H2D raises through T_H2C and T_H2B to T_H2A's handler, which then calls T_Quiet.
     Calls=@(
        @{ Function='T_H2A'; Depth='1'; Parent=-1; Caller='none'; Outcome='handled' }
        @{ Function='T_H2B'; Depth='2'; Parent=0; Outcome='unwound' }
        @{ Function='T_H2C'; Depth='3'; Parent=1; Outcome='unwound' }
        @{ Function='T_H2D'; Depth='4'; Parent=2; Outcome='threw' }
        @{ Function='T_Quiet'; Depth='2'; Parent=0; Outcome='returned' } )
     Why='THREE frames released at once, then execution continues in the
          handler -- the case most likely to leave the stack drifted' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
