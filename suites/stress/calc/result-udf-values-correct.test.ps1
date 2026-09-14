$case = @{ Name='result-udf-values-correct'
     Modules=@{
       'M'=@'
Public Function Precise(ByVal n As Double) As Double
    Precise = n * 1.5 + 0.25
End Function
'@
     }
     Cells=@{ 'A1'='=Precise(1)'; 'A2'='=Precise(2)'; 'A3'='=Precise(4)'; 'A4'='=Precise(8)'; 'A5'='=Precise(16)' }
     FillFormula=@{ Range='B1:B20'; Formula='=Precise(ROW())' }
     VerifyCells=@{ 'A1'='1.75'; 'A2'='3.25'; 'A3'='6.25'; 'A4'='12.25'; 'A5'='24.25' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        if ($t.framesOpened -lt 10) { return "expected >=10 activations, got $($t.framesOpened)" }
        $null }
     Why='a traced UDF across many cells must still put the RIGHT numbers in them' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
