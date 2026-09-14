$case = @{ Name='loop-exact'
     Setup=@'
Public Sub T_Loop(ByVal n As Long)
    Dim i As Long
    Dim s As Double
    For i = 1 To n
        s = s + 1
    Next i
End Sub
'@
     Invoke=@{ Name='T_Loop'; Args=@(10000) }
     Expect={ param($t)
        # one body statement + the Next => 2 BoS per iteration
        $lo = 2 * 10000; $hi = 2 * 10000 + 20
        if ($t.statements -lt $lo -or $t.statements -gt $hi) {
            return "expected ~$lo statements for 10000 iterations, got $($t.statements)" }
        $null }
     Why='THE load-bearing arithmetic: VBA semantics predict the count exactly' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
