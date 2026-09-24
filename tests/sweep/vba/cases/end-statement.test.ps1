$case = @{ Name='end-statement'
     Setup=@'
Public Sub T_EndC()
    End
End Sub
Public Sub T_EndB()
    Call T_EndC
End Sub
Public Sub T_EndA()
    Call T_EndB
End Sub
Public Sub T_AfterEnd()
    Call T_AfterEnd2
End Sub
Public Sub T_AfterEnd2()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Name='T_EndA'; Args=@(); MayRaise=$true }
     Then=@{ Name='T_AfterEnd' }
     Expect={ param($t)
        # `End` tears the session down with no handler or exit opcode, so the follow-up call is
        # the real test: unclosed frames would push its clean chain deeper.
        if ($t.faults -gt 0) { return "$($t.faults) faults across End" }
        if ($t.maxDepth -gt 4) {
            return "DRIFT after End: depth reached $($t.maxDepth); nothing here nests past 3" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK across End: opened $($t.framesOpened), closed $($t.framesClosed)" }

        # The only source of `abandoned`: a user break is trappable error 18, and an XLL is only
        # asked to stop, so End is the single trigger.
        $ended = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.function -like 'T_End*' })
        if ($ended.Count -lt 3) { return "expected 3 killed frames, saw $($ended.Count)" }
        $notAbandoned = @($ended | Where-Object { $_.outcome -ne 'abandoned' })
        if ($notAbandoned.Count) {
            return ("End killed frames that do not read 'abandoned': " +
                    (($notAbandoned | ForEach-Object { "$($_.function)=$($_.outcome)" }) -join ',')) }

        # the later, clean call must not have caught it
        $after = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.function -like 'T_AfterEnd*' })
        $wrong = @($after | Where-Object { $_.outcome -ne 'returned' })
        if ($wrong.Count) {
            return ("a clean call after End does not read 'returned': " +
                    (($wrong | ForEach-Object { "$($_.function)=$($_.outcome)" }) -join ',')) }
        $null }
     # End kills all three frames; the later call is a fresh chain from depth 1, not nested under them.
     Calls=@(
        @{ Function='T_EndA'; Depth='1'; Parent=-1; Caller='none'; Outcome='abandoned' }
        @{ Function='T_EndB'; Depth='2'; Parent=0; Outcome='abandoned' }
        @{ Function='T_EndC'; Depth='3'; Parent=1; Outcome='abandoned' }
        @{ Function='T_AfterEnd';  Depth='1'; Parent=-1; Caller='none'; Outcome='returned' }
        @{ Function='T_AfterEnd2'; Depth='2'; Parent=3; Outcome='returned' } )
     Why='`End` abandons frames rather than unwinding them. Does a later call
          in the SAME session inherit them?' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
