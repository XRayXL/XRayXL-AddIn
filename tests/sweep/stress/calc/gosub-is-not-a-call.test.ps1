# GoSub moves rsp within one activation; it must not open a frame.
# Excel picks its own pass count, so assert GoSubbed == Plain (same shape, no GoSub).
# Recursed must still nest: it guards against ignoring rsp changes altogether.
$case = @{ Name='gosub-is-not-a-call'
     Modules=@{
       'M'=@'
Public Function GoSubbed(ByVal n As Double) As Double
    GoSub Helper
    GoSubbed = n
    Exit Function
Helper:
    n = n + 1
    Return
End Function
Public Function Plain(ByVal n As Double) As Double
    Dim m As Double
    m = n + 1
    Plain = n
End Function
Public Function Recursed(ByVal n As Double) As Double
    If n <= 0 Then
        Recursed = 0
    Else
        Recursed = n + Recursed(n - 1)
    End If
End Function
'@
     }
     Cells=@{
       'A1'='=GoSubbed(1)'; 'A2'='=GoSubbed(2)'; 'A3'='=GoSubbed(3)'
       'C1'='=Plain(1)';    'C2'='=Plain(2)';    'C3'='=Plain(3)'
       'E1'='=Recursed(3)'
     }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $entries = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
        $gs = @($entries | Where-Object { $_.function -eq 'GoSubbed' }).Count
        $pl = @($entries | Where-Object { $_.function -eq 'Plain'    }).Count
        $rc = @($entries | Where-Object { $_.function -eq 'Recursed' }).Count

        # Inventing a frame and closing it balances too, so this only rules out
        # leaking; the checks below are about counting.
        if ($t.framesOpened -ne $t.framesClosed) {
            return "frames leaked: $($t.framesOpened)/$($t.framesClosed)" }

        # Control: Plain must have been called at all, or there is nothing to compare against.
        if ($pl -lt 3) {
            return "CONTROL FAILED: Plain called from 3 cells but only $pl activation(s) traced" }

        # Control: a real nested call at a different rsp must still nest. Recursed(3) is 4 activations.
        if ($rc -lt 4) {
            return "CONTROL FAILED: Recursed(3) should nest 4 deep but only $rc activation(s) traced -- a genuine nested call is no longer being opened" }
        if ($t.maxDepth -lt 4) {
            return "CONTROL FAILED: maxDepth $($t.maxDepth), expected >= 4 from Recursed(3)" }

        # Identical shape and cell count, differing only by the GoSub: anything but equality means
        # rsp is still opening frames.
        if ($gs -ne $pl) {
            return "GoSubbed traced $gs times and Plain $pl from the same number of cells -- the GoSub is being counted as a nested call" }
        $null }
     Why='GoSub moves the interpreter rsp inside one activation; frames must be
          opened by the p-code prologue alone, or a GoSub is counted as a nested
          call -- paired with an identical GoSub-free procedure so the count is
          comparable however often Excel recalculates, and with a recursion
          control which reaches the same procedure at a different rsp and MUST
          still nest' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
