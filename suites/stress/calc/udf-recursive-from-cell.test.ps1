$case = @{ Name='udf-recursive-from-cell'
     Modules=@{
       'M'=@'
Public Function RecCell(ByVal n As Double) As Double
    If n <= 0 Then
        RecCell = 0
    Else
        RecCell = n + RecCell(n - 1)
    End If
End Function
'@
     }
     Cells=@{ 'A1'='=RecCell(40)' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        if ($t.recursions -lt 39) { return "expected >=39 recursions, got $($t.recursions)" }
        if ($t.maxDepth -lt 40)   { return "expected depth >=40, got $($t.maxDepth)" }
        $null }
     Why='a UDF that recurses 40 deep from a cell: recursion inside the calc engine' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
