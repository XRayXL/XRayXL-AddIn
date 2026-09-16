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
     # The Calc trigger dirties the sheet and then rebuilds, which evaluates A1 twice (measured,
     # see end-statement-midchain). Each pass enters BadTop from the cell, so it reads unhandled.
     Calls=@(0, 3 | ForEach-Object {
        @{ Function='BadTop';  Depth='1'; Parent=-1; Caller='cell'; Cell='A1'; Outcome='unhandled' }
        @{ Function='BadMid';  Depth='2'; Parent=$_;       Outcome='unwound' }
        @{ Function='BadLeaf'; Depth='3'; Parent=($_ + 1); Outcome='threw' } })
     Why='an unhandled error inside a cell-invoked UDF: it unwinds with NO exit opcode. NOTE the names -- Bad1/Bad3 are CELL ADDRESSES (BAD is a legal column), so a UDF so named yields #REF! and never runs' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
