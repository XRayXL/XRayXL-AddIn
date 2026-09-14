$case = @{ Name='deep-recursion-100'
     Setup=@'
Public Sub T_Deep(ByVal n As Long)
    If n > 0 Then
        Call T_Deep(n - 1)
    End If
End Sub
'@
     Invoke=@{ Name='T_Deep'; Args=@(100) }
     Expect={ param($t)
        if ($t.overflows -gt 0)    { return "shadow stack overflowed at depth 101 (capacity 256)" }
        if ($t.faults -gt 0)       { return "$($t.faults) guarded reads faulted at depth 100" }
        if ($t.recursions -ne 100) { return "expected exactly 100 recursions, got $($t.recursions)" }
        if ($t.maxDepth -ne 101)   { return "expected depth 101, got $($t.maxDepth)" }
        $null }
     Why='100 levels deep, counted exactly and without overflowing' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
