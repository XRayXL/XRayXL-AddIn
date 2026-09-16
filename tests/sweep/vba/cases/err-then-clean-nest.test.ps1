$case = @{ Name='err-then-clean-nest'
     Setup=@'
Public Sub T_DriftRaise()
    Err.Raise 5, "T_DriftRaise", "deliberate"
End Sub
Public Sub T_DriftMid()
    Call T_DriftRaise
End Sub
Public Sub T_DriftTop()
    On Error Resume Next
    Call T_DriftMid
    On Error GoTo 0
    Call T_DriftN1
End Sub
Public Sub T_DriftN1()
    Call T_DriftN2
End Sub
Public Sub T_DriftN2()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Name='T_DriftTop'; Args=@() }
     Expect={ param($t)
        # THE DRIFT DETECTOR. After the unwind the same session runs a call
        # chain of known shape. If the unwound frames were never popped, the
        # later chain sits on top of them and the depth comes out too big.
        # Deepest legitimate point is T_DriftTop > T_DriftMid > T_DriftRaise = 3.
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        if ($t.maxDepth -gt 4) {
            return "DRIFT: depth reached $($t.maxDepth); nothing here nests deeper than 3-4" }
        if ($t.faults -gt 0) { return "$($t.faults) faults" }
        $null }
     # The later chain must sit at depths 2 and 3 under T_DriftTop, not on stale frames.
     Calls=@(
        @{ Function='T_DriftTop'; Depth='1'; Parent=-1; Caller='none'; Outcome='handled' }
        @{ Function='T_DriftMid'; Depth='2'; Parent=0; Outcome='unwound' }
        @{ Function='T_DriftRaise'; Depth='3'; Parent=1; Outcome='threw' }
        @{ Function='T_DriftN1'; Depth='2'; Parent=0; Outcome='returned' }
        @{ Function='T_DriftN2'; Depth='3'; Parent=3; Outcome='returned' } )
     Why='does tracing stay CORRECT after an unwind? Run a known-shape call
          chain afterwards and check the depth has not inherited stale frames' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
