# `End` from a cell UDF leaves no stale frames. Unclosed frames only show in the next call, so
# the follow-up chain is what is asserted.
$case = @{ Name='end-from-udf'
     Setup=@'
Public Function T_EndF() As Double
    T_EndF = T_EndF2()
End Function
Private Function T_EndF2() As Double
    End
End Function
Public Sub T_AfterEndF()
    Call T_AfterEndF2
End Sub
Public Sub T_AfterEndF2()
    Dim a As Long
    a = 1
End Sub
'@
     Invoke=@{ Formula='=T_EndF()' }
     Then=@{ Name='T_AfterEndF' }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) faults" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }

        # `end-statement` pins the killed frames' outcome. Write nothing to the pipeline here:
        # any output is taken as the failure reason.
        $ended = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.function -like 'T_EndF*' })
        $killed = "killed by End: " +
                  ((($ended | ForEach-Object { "$($_.function)=$($_.outcome)" }) -join ','))

        # deeper than 2 means the torn-down frames were still on the shadow stack
        $after = @($t.rows | Where-Object { $_.kind -eq 'exit' -and $_.function -like 'T_AfterEndF*' })
        if ($after.Count -lt 2) {
            return "the follow-up chain did not run: $($after.Count) of 2 frames closed ($killed)" }
        $wrong = @($after | Where-Object { $_.outcome -ne 'returned' })
        if ($wrong.Count) {
            return ("a clean call after End does not read 'returned': " +
                    (($wrong | ForEach-Object { "$($_.function)=$($_.outcome)" }) -join ',') +
                    " ($killed)") }
        $deepest = ($after | ForEach-Object { [int]$_.depth } | Sort-Object)[-1]
        if ($deepest -gt 2) {
            return ("DRIFT after End: the follow-up chain reached depth $deepest;" +
                    " nothing here nests past 2 ($killed)") }
        $null }
     Why='the same teardown reached from a cell rather than a macro' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
