$case = @{ Name='sibling-calls'
     Setup=@'
Public Sub T_S1()
    Dim a As Long
    a = 1
End Sub
Public Sub T_S2()
    Dim a As Long
    a = 1
End Sub
Public Sub T_Sib()
    Call T_S1
    Call T_S2
    Call T_S1
End Sub
'@
     Invoke=@{ Name='T_Sib'; Args=@() }
     Expect={ param($t)
        if ($t.procedures -lt 3)  { return "expected 3 procedures, got $($t.procedures)" }
        if ($t.maxDepth -ne 2)    { return "siblings should nest only 2 deep, got $($t.maxDepth)" }
        # A RETURN is only observed if the caller executes another statement
        # before its next call -- here it does not, so the caller is not seen
        # again between siblings. Transitions therefore count 4, not 6. The
        # exit slots are the reliable return signal; transitions are not.
        if ($t.transitions -lt 3) { return "expected >=3 transitions, got $($t.transitions)" }
        if ($t.exits -lt 3)       { return "expected >=3 exits for 3 calls, got $($t.exits)" }
        $null }
     Why='returns are only visible via the EXIT slots -- a caller that calls
          again immediately is never observed in between' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
