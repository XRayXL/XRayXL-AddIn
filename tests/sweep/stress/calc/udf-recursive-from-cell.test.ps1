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
     # The Calc trigger evaluates A1 twice (measured, see end-statement-midchain). Each pass is
     # RecCell(40) down to RecCell(0), 41 activations, and RecCell(n) returns n(n+1)/2.
     Calls=@(0, 41 | ForEach-Object { $base = $_
        0..40 | ForEach-Object { $n = 40 - $_
            $call = @{ Function='RecCell'; Args="a1:Double=$n"; Ret="$($n * ($n + 1) / 2)"; RetType='Double'
                       Depth="$($_ + 1)"; Parent=$(if ($_ -eq 0) { -1 } else { $base + $_ - 1 }); Outcome='returned' }
            if ($_ -eq 0) { $call.Caller = 'cell'; $call.Cell = 'A1' }
            $call } })
     Why='a UDF that recurses 40 deep from a cell: recursion inside the calc engine' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
