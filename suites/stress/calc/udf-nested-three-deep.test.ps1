$case = @{ Name='udf-nested-three-deep'
     Modules=@{
       'M'=@'
Public Function NestA(ByVal x As Double) As Double
    NestA = NestB(x) + 1
End Function
Public Function NestB(ByVal x As Double) As Double
    NestB = NestC(x) + 1
End Function
Public Function NestC(ByVal x As Double) As Double
    NestC = x * 2
End Function
'@
     }
     FillFormula=@{ Range='A1:A50'; Formula='=NestA(ROW())' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'NestA','NestB','NestC'; if ($why) { return $why }
        if ($t.maxDepth -lt 3) { return "expected depth >=3, got $($t.maxDepth)" }
        $null }
     Why='UDFs calling UDFs three levels down, from many cells at once' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
