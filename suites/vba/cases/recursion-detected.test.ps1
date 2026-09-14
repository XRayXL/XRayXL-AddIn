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
        # T_Rec(8) makes 8 recursive calls, so 9 activations of ONE procedure.
        # The trailer alone cannot see this -- every activation shares it. The
        # interpreter's rsp at the dispatch falls one step per call, so
        # (trailer, sp) identifies the ACTIVATION and the count is exact.
        if ($t.procedures -ne 1) { return "expected exactly 1 procedure, got $($t.procedures)" }
        if ($t.recursions -ne 8) { return "expected exactly 8 recursions, got $($t.recursions)" }
        if ($t.maxDepth   -ne 9) { return "expected depth 9 (9 activations), got $($t.maxDepth)" }
        $null }
     Why='RECURSION, exactly counted. The trailer names a procedure; the stack
          pointer names an activation of it' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
