$case = @{ Name='same-name-two-modules'
     Modules=@{
       'ModA'=@'
Public Sub Common()
    Dim z As Long
    z = 1
End Sub
Public Sub CallA()
    ModA.Common
End Sub
'@
       'ModB'=@'
Public Sub Common()
    Dim z As Long
    z = 2
End Sub
Public Sub CallB()
    ModB.Common
End Sub
'@
       'M'=@'
Public Sub Go()
    ModA.CallA
    ModB.CallB
End Sub
'@
     }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $rows = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'Common' })
        if ($rows.Count -lt 2) { return "expected both Common procedures, got $($rows.Count)" }
        $procs = @($rows | ForEach-Object { $_.proc } | Sort-Object -Unique)
        if ($procs.Count -lt 2) {
            return "both Common procedures reported the SAME trailer -- identity collapsed" }
        $mods = @($rows | ForEach-Object { $_.module } | Sort-Object -Unique)
        if ($mods.Count -lt 2) { return "both reported the same module: $($mods -join ',')" }
        $null }
     Why='two modules exporting the SAME procedure name: they must not collapse into one' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
