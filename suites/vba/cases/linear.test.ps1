$case = @{ Name='linear'
     Setup=@'
Public Sub T_Linear()
    Dim a As Long
    a = 1
    a = 2
    a = 3
    a = 4
    a = 5
End Sub
'@
     Invoke=@{ Name='T_Linear'; Args=@() }
     Expect={ param($t)
        # five assignments and End Sub; an uninitialised Dim may or may not carry a statement marker
        if ($t.statements -lt 6 -or $t.statements -gt 7) { return "expected 6 or 7 statements, got $($t.statements)" }
        if ($t.maxDepth -ne 1)    { return "depth should be 1 for a flat procedure, got $($t.maxDepth)" }
        $null }
     Why='statement count tracks source statements, and a flat call has depth 1' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
