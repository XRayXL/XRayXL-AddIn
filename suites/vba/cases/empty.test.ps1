$case = @{ Name='empty'
     Setup=@'
Public Sub T_Empty()
End Sub
'@
     Invoke=@{ Name='T_Empty'; Args=@() }
     Expect={ param($t)
        if ($t.statements -lt 1)   { return "no statements counted for a procedure that ran" }
        if ($t.procedures -lt 1)   { return "no procedure recorded" }
        if ($t.faults -gt 0)       { return "$($t.faults) guarded reads faulted" }
        $null }
     Why='the smallest possible procedure still registers as one' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
