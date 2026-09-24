$case = @{ Name='recursion-detected'
     Setup=@'
Public Sub T_Rec(ByVal n As Long)
    If n > 0 Then
        Call T_Rec(n - 1)
    End If
End Sub
'@
     Invoke=@{ Name='T_Rec'; Args=@(8) }
     Expect={ param($t)
        # every activation shares the trailer; the interpreter's rsp is what tells them apart
        if ($t.procedures -ne 1) { return "expected exactly 1 procedure, got $($t.procedures)" }
        if ($t.recursions -ne 8) { return "expected exactly 8 recursions, got $($t.recursions)" }
        if ($t.maxDepth   -ne 9) { return "expected depth 9 (9 activations), got $($t.maxDepth)" }
        $null }
     Calls=@(0..8 | ForEach-Object { @{ Function='T_Rec'; Args="a1:Long=$(8 - $_)"; Depth="$($_ + 1)"; Parent=($_ - 1); Outcome='returned' } })
     Why='RECURSION, exactly counted. The trailer names a procedure; the stack
          pointer names an activation of it' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
