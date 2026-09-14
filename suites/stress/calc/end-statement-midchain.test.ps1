$case = @{ Name='end-statement-midchain'
     Modules=@{
       'M'=@'
Public Sub E3()
    End
End Sub
Public Sub E2()
    E3
End Sub
Public Sub E1()
    E2
End Sub
Public Function Kick() As Double
    E1
    Kick = 1
End Function
'@
     }
     Cells=@{ 'A1'='=Kick()' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        # `End` tears the VBA session down and fires NO exit opcode, so the
        # frames it abandons cannot close themselves. Leaving them to the
        # stack-pointer backstop does not work: it needs a later statement at a
        # HIGHER rsp, and a full rebuild evaluates this sheet TWICE with the
        # second pass starting DEEPER, so the new activations nest underneath
        # the dead ones. Measured that way: depth 8 where 4 is right, frames
        # 8/8.
        #
        # The End opcode itself (slot 619) is now hooked and closes the chain it
        # kills, so 4 is asserted rather than tolerated. THE OLD NUMBER IS THE
        # REGRESSION SIGNAL: 8 means the End hook stopped firing and the
        # backstop is carrying it again.
        if ($t.maxDepth -ne 4) {
            return "depth $($t.maxDepth), expected 4" +
                   $(if ($t.maxDepth -eq 8) { ' -- 8 is the pre-fix value: the End hook is not firing' }) }
        # The safety property, unchanged and still checked: nothing leaks.
        if ($t.framesOpened -ne $t.framesClosed) {
            return "frames leaked across End: $($t.framesOpened)/$($t.framesClosed)" }
        $null }
     Why='End tears the session down three frames deep, firing no exit opcodes at all. Asserts that the End hook closes the chain it kills, so a rebuild starting deeper does not nest under dead frames' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
