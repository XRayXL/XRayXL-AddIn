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
     # Go, then A 30, B 29, C 28, A 27 ... down to 0: 31 activations, each inside the last.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(0..30 | ForEach-Object { @{ Function=@('A','B','C')[$_ % 3]; Args="a1:Long=$(30 - $_)"; Depth="$($_ + 2)"; Parent=$_; Outcome='returned' } })
     Why='three procedures recursing through each other, so the trailer changes every frame' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
