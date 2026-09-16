$case = @{ Name='dep-chain-three-workbooks'
     Modules=@{
       'M'=@'
Public Sub ChainA()
    Application.Run "DepE.xlsm!ChainB"
End Sub
'@
     }
     Deps=@(
       @{ Name='DepE'; Modules=@{
         'ModE'=@'
Public Sub ChainB()
    ChainC
End Sub
Public Sub ChainC()
    Dim z As Long
    z = 3
End Sub
'@
       } }
     )
     Trigger=@{ Kind='Run'; Name='ChainA' }
     Expect={ param($t)
        $why = Assert-VbaTraced $t 'ChainA','ChainB','ChainC'; if ($why) { return $why }
        $mods = @($t.rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') } | ForEach-Object { $_.module } | Sort-Object -Unique)
        if ($mods.Count -lt 2) { return "expected 2 distinct workbooks, saw: $($mods -join ',')" }
        if ($t.maxDepth -lt 3) { return "expected depth >=3 across books, got $($t.maxDepth)" }
        $null }
     Why='A crosses into another workbook and goes two deeper there: one stack, two books, three frames' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
