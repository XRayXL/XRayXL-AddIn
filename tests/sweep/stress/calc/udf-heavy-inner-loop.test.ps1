$case = @{ Name='udf-heavy-inner-loop'
     Modules=@{
       'M'=@'
Public Function Heavy(ByVal n As Double) As Double
    Dim i As Long
    Dim s As Double
    For i = 1 To 20000
        s = s + i
    Next i
    Heavy = s + n
End Function
'@
     }
     FillFormula=@{ Range='A1:A20'; Formula='=Heavy(ROW())' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        if ($t.statements -lt 400000) { return "expected >=400k statements, got $($t.statements)" }
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted under load" }
        $null }
     # Each cell is evaluated twice by the Calc trigger, in an order Excel decides. The loop
     # sums 1..20000 = 200010000, so Heavy(n) returns 200010000 + n.
     CallsAnyOrder=$true
     Calls=@(1..2 | ForEach-Object { 1..20 | ForEach-Object {
        @{ Function='Heavy'; Args="a1:Double=$_"; Ret="$(200010000 + $_)"; RetType='Double'
           Depth='1'; Caller='cell'; Cell="A$_"; Outcome='returned' } } })
     Why='a UDF with a 20,000-iteration inner loop in each of 20 cells: 400k statements' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
