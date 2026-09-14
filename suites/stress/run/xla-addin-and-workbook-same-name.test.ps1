$case = @{ Name='xla-addin-and-workbook-same-name'
     Modules=@{
       'M'=@'
Public Sub Helper()
    Dim z As Long
    z = 2
End Sub
Public Sub Go()
    Helper
    Application.Run "CallAddinHelper"
End Sub
'@
     }
     Deps=@(
       @{ Name='AddThree'; IsAddin=$true; Modules=@{
         'MA'=@'
Public Sub Helper()
    Dim z As Long
    z = 1
End Sub
Public Sub CallAddinHelper()
    Helper
End Sub
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $rows = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Helper' })
        if ($rows.Count -lt 2) { return "expected Helper from both, got $($rows.Count)" }
        $procs = @($rows | ForEach-Object { $_.proc } | Sort-Object -Unique)
        if ($procs.Count -lt 2) { return "add-in and workbook Helper shared ONE trailer" }
        $mods = @($rows | ForEach-Object { $_.module } | Sort-Object -Unique)
        if ($mods.Count -lt 2) { return "both attributed to: $($mods -join ',')" }
        $null }
     Why='the add-in AND the workbook both define Helper(): they must stay distinct' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
