$case = @{ Name='mutual-recursion-3way'
     Modules=@{
       'M'=@'
Public Sub A(ByVal n As Long)
    If n > 0 Then B n - 1
End Sub
Public Sub B(ByVal n As Long)
    If n > 0 Then C n - 1
End Sub
Public Sub C(ByVal n As Long)
    If n > 0 Then A n - 1
End Sub
Public Sub Go()
    A 30
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        if ($t.procedures -lt 4) { return "expected 4 procedures, got $($t.procedures)" }
        if ($t.maxDepth -lt 9)    { return "expected depth >=9, got $($t.maxDepth)" }
        $null }
     Why='three procedures recursing through each other, so the trailer changes every frame' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
