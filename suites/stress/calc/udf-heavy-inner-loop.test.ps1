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
     Why='a UDF with a 20,000-iteration inner loop in each of 20 cells: 400k statements' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
