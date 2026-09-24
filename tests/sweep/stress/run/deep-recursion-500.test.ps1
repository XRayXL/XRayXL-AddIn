$case = @{ Name='deep-recursion-500'
     Modules=@{
       'M'=@'
Public Sub Deep(ByVal n As Long)
    If n > 0 Then Deep n - 1
End Sub
Public Sub Go()
    Deep 500
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        # VBA goes deeper than the shadow stack holds: capping is allowed, reporting the cap as the
        # depth is not.
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted at depth" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "frames leaked: $($t.framesOpened)/$($t.framesClosed)" }

        if ($t.maxDepth -ge 256) {
            # It capped. Then it must say so, and still know the real depth.
            if ($t.overflows -lt 1) {
                return "hit the 256-frame cap but counted NO overflows -- frames vanished silently" }
            # Every activation is either recorded or counted as an overflow;
            # ipEntries is the independent total from the p-code boundary.
            if ($t.framesOpened + $t.overflows -ne $t.ipEntries) {
                return "activations unaccounted: framesOpened $($t.framesOpened) + overflows $($t.overflows) != ipEntries $($t.ipEntries)" }
            # deepestSeen is a floor on how deep VBA went; 256 would mean the cap is reported as the answer.
            if ($t.deepestSeen -le $t.maxDepth) {
                return "deepestSeen $($t.deepestSeen) is no deeper than the cap $($t.maxDepth) -- the true depth is being lost" }
            if ($t.deepestSeen -lt 500) {
                return "VBA went 502 activations deep but deepestSeen is only $($t.deepestSeen)" }
        } elseif ($t.recursions -lt 499) {
            return "did not cap, but only $($t.recursions) recursions of 499" }
        $null }
     # The shadow stack holds 256 frames (kMaxDepth): Go, then Deep 500 down to Deep 246 are written.
     # Deeper activations are only counted, and one depth-capped row names the deepest written call.
     Calls=@(@{ Function='Go'; Depth='1'; Parent=-1; Outcome='returned' }) +
           @(0..254 | ForEach-Object { @{ Function='Deep'; Args="a1:Long=$(500 - $_)"; Depth="$($_ + 2)"; Parent=$_; Outcome='returned' } })
     DepthCappedUnder=255
     Why='recursion far past any sane VBA stack. Frames may be capped, but the
          DEPTH must not be: reporting maxDepth=256 for a 502-deep recursion is
          a wrong number, not a truncated one, so deepestSeen must exceed the
          cap and every activation must be either recorded or counted' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
