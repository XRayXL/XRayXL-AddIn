$case = @{ Name='exits-match-frames'
     Setup=@'
Public Sub T_X3()
    Dim a As Long
    a = 1
End Sub
Public Sub T_X2()
    Call T_X3
End Sub
Public Sub T_X1()
    Call T_X2
    Call T_X2
End Sub
'@
     Invoke=@{ Name='T_X1'; Args=@() }
     Expect={ param($t)
        # Two independent mechanisms count the same thing: the exit opcodes
        # fire once per procedure exit, and the shadow stack closes one frame
        # per return. They are computed from different signals -- opcode slots
        # versus the stack pointer -- so agreement is real evidence.
        if ($t.exits -lt 5) { return "expected >=5 exit opcodes, got $($t.exits)" }
        $closed = $t.framesClosed
        if ($closed -lt ($t.exits - 2) -or $closed -gt ($t.exits + 2)) {
            return "framesClosed=$closed disagrees with exits=$($t.exits)" }
        $null }
     Why='CROSS-CHECK: exits come from the opcode slots, frame closes come from
          the stack pointer. Two unrelated signals must agree' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
