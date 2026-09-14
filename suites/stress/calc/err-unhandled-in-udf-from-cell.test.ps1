$case = @{ Name='err-unhandled-in-udf-from-cell'
     Modules=@{
       'M'=@'
Public Function BadTop() As Double
    BadTop = BadMid()
End Function
Public Function BadMid() As Double
    BadMid = BadLeaf()
End Function
Public Function BadLeaf() As Double
    Err.Raise 5301
End Function
'@
     }
     Cells=@{ 'A1'='=BadTop()' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        if ($t.framesOpened -lt 3) { return "expected >=3 frames, got $($t.framesOpened)" }
        $null }
     Why='an unhandled error inside a cell-invoked UDF: it unwinds with NO exit opcode. NOTE the names -- Bad1/Bad3 are CELL ADDRESSES (BAD is a legal column), so a UDF so named yields #REF! and never runs' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
