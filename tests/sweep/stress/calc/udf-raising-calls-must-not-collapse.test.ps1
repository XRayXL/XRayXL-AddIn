# A cell UDF that raises never reaches its last epilogue, so its frame can stay
# open and swallow the next call. Boom and Fine differ only in raising: if both
# fail, suspect the case, not the tracer.
$case = @{ Name='udf-raising-calls-must-not-collapse'
     Modules=@{
       'M'=@'
Public Function Boom(ByVal n As Double) As Double
    Err.Raise 5401
End Function
Public Function Fine(ByVal n As Double) As Double
    Fine = n * 2
End Function
'@
     }
     Cells=@{
       'A1'='=Boom(1)'; 'A2'='=Boom(2)'; 'A3'='=Boom(3)'
       'A4'='=Boom(4)'; 'A5'='=Boom(5)'; 'A6'='=Boom(6)'
       'C1'='=Fine(1)'; 'C2'='=Fine(2)'; 'C3'='=Fine(3)'
       'C4'='=Fine(4)'; 'C5'='=Fine(5)'; 'C6'='=Fine(6)'
     }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $entries = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
        $boom = @($entries | Where-Object { $_.function -eq 'Boom' }).Count
        $fine = @($entries | Where-Object { $_.function -eq 'Fine' }).Count

        # The control: six cells of a procedure that returns normally. If this is short, the problem
        # is not the raise.
        if ($fine -lt 6) {
            return "CONTROL FAILED: Fine returns normally from 6 cells but only $fine activation(s) traced -- the collapse is not specific to raising, or the case is wrong" }

        # The defect: every call raises, so no epilogue closes the frame.
        if ($boom -lt 6) {
            return "COLLAPSE: Boom was called from 6 cells but only $boom activation(s) traced -- a raising UDF leaves its frame open and swallows the next call" }

        # A collapsed activation keeps whichever argument was captured, so the row would assert a
        # call that did not happen as written.
        $args = @($entries | Where-Object { $_.function -eq 'Boom' } | ForEach-Object { $_.args })
        $distinct = @($args | Sort-Object -Unique).Count
        if ($distinct -lt 6) {
            return "Boom traced $boom times but only $distinct distinct argument value(s) -- activations were merged" }
        $null }
     Why='a cell-invoked UDF that raises never reaches its statement epilogue,
          so nothing closes its frame and the next call at the same rsp is
          mistaken for a continuation -- paired with a non-raising control of
          identical shape so a failure cannot be blamed on the harness' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
