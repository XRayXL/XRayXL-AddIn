$case = @{ Name='dep-same-proc-name-two-workbooks'
     Modules=@{
       'M'=@'
Public Function Calc(ByVal n As Double) As Double
    Calc = n + 1
End Function
'@
     }
     Cells=@{ 'A1'='=Calc(1)'; 'A2'='=''DepC.xlsm''!Calc(2)' }
     Deps=@(
       @{ Name='DepC'; Modules=@{
         'MC'=@'
Public Function Calc(ByVal n As Double) As Double
    Calc = n + 100
End Function
'@
       } }
     )
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $rows = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Calc' })
        if ($rows.Count -lt 2) { return "expected Calc from both workbooks, got $($rows.Count)" }
        $procs = @($rows | ForEach-Object { $_.proc } | Sort-Object -Unique)
        if ($procs.Count -lt 2) { return "both Calc procedures shared ONE trailer -- identity collapsed" }
        $mods = @($rows | ForEach-Object { $_.module } | Sort-Object -Unique)
        if ($mods.Count -lt 2) { return "both attributed to the same workbook: $($mods -join ',')" }
        $null }
     Why='both workbooks define Calc(): two books, two trailers, two module labels' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
