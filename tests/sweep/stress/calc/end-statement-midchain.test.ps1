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
        # `End` tears the VBA session down and fires no exit opcode, so the frames it abandons
        # cannot close themselves, and the stack-pointer backstop cannot do it either: a full
        # rebuild evaluates this sheet twice with the second pass starting deeper, so the new
        # activations would nest underneath the dead ones. The End opcode (slot 619) is hooked
        # and closes the chain it kills. A depth of 8 here means that hook stopped firing.
        if ($t.maxDepth -ne 4) {
            return "depth $($t.maxDepth), expected 4" +
                   $(if ($t.maxDepth -eq 8) { ' -- 8 is the pre-fix value: the End hook is not firing' }) }
        # The safety property, unchanged and still checked: nothing leaks.
        if ($t.framesOpened -ne $t.framesClosed) {
            return "frames leaked across End: $($t.framesOpened)/$($t.framesClosed)" }
        $null }
     # Two passes (see above), each a fresh chain from depth 1: End kills all four frames, so all read abandoned.
     Calls=@(0, 4 | ForEach-Object {
        @{ Function='Kick'; Depth='1'; Parent=-1; Caller='cell'; Cell='A1'; Outcome='abandoned' }
        @{ Function='E1';   Depth='2'; Parent=$_;       Outcome='abandoned' }
        @{ Function='E2';   Depth='3'; Parent=($_ + 1); Outcome='abandoned' }
        @{ Function='E3';   Depth='4'; Parent=($_ + 2); Outcome='abandoned' } })
     Why='End tears the session down three frames deep, firing no exit opcodes at all. Asserts that the End hook closes the chain it kills, so a rebuild starting deeper does not nest under dead frames' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
