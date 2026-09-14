$case = @{ Name='udf-200-cells'
     Modules=@{
       'M'=@'
Public Function Cell200(ByVal n As Double) As Double
    Cell200 = n * 2
End Function
'@
     }
     FillFormula=@{ Range='A1:A200'; Formula='=Cell200(ROW())' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $entries = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Cell200' })
        # at least 200: how often a rebuild calls a UDF beyond once per cell is Excel's business
        if ($entries.Count -lt 200) { return "expected >=200 activations, got $($entries.Count)" }
        $cells = @($entries | ForEach-Object { Get-CallerCell $_ } | Where-Object { $_ } | Sort-Object -Unique)
        if ($cells.Count -ne 200) { return "calls came from $($cells.Count) distinct cells, expected all of A1:A200" }
        # a VBA UDF is never thread-safe, so Excel runs every call on its main thread
        $threads = @($entries | ForEach-Object { $_.thread } | Sort-Object -Unique)
        if ($threads.Count -ne 1) { return "Cell200 ran on $($threads.Count) threads: $($threads -join ',')" }
        $null }
     Why='one UDF across 200 cells: an activation from every cell, all on the one thread Excel runs VBA on' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
