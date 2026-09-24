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
        if ($t.faults -gt 0)       { return "$($t.faults) guarded reads faulted at depth 100" }
        if ($t.recursions -ne 100) { return "expected exactly 100 recursions, got $($t.recursions)" }
        if ($t.maxDepth -ne 101)   { return "expected depth 101, got $($t.maxDepth)" }
        $null }
     # T_Deep(100) down to T_Deep(0): 101 activations, each inside the one before.
     Calls=@(0..100 | ForEach-Object { @{ Function='T_Deep'; Args="a1:Long=$(100 - $_)"; Depth="$($_ + 1)"; Parent=($_ - 1); Outcome='returned' } })
     Why='100 levels deep, counted exactly' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
