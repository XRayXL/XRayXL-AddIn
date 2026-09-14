$case = @{ Name='nesting-3'
     Setup=@'
Public Sub T_N1()
    Call T_N2
End Sub
Public Sub T_N2()
    Call T_N3
End Sub
Public Sub T_N3()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Name='T_N1'; Args=@() }
     Expect={ param($t)
        if ($t.maxDepth -lt 3)  { return "expected depth 3 for a 3-deep call chain, got $($t.maxDepth)" }
        if ($t.procedures -lt 3){ return "expected 3 distinct procedures, got $($t.procedures)" }
        if ($t.recursions -ne 0){ return "reported $($t.recursions) recursions where there are none" }
        # PLANTED ANSWER: we know what these procedures are called, so the
        # tracer has to agree. "A name appeared" is not the same claim.
        if ($t.unnamed -ne 0)   { return "$($t.unnamed) procedure(s) could not be named" }
        foreach ($want in @('T_N1','T_N2','T_N3')) {
            if ($t.names -notcontains $want) {
                return "name '$want' missing; resolved: $($t.names -join ',')" }
        }
        $null }
     Why='nesting AND identity: three known procedures, resolved to the names
          they were written with' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
