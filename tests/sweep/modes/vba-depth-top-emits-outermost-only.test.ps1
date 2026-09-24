# VBA DEPTH=TOP emits rows for the depth-1 frame of a cell-called chain alone, while the totals
# still count every frame: the trace thins but the accounting does not.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Public Function T_L1() As Double
    T_L1 = T_L2() + 1
End Function
Private Function T_L2() As Double
    T_L2 = T_L3() + 2
End Function
Private Function T_L3() As Double
    T_L3 = 4
End Function
'@
    New-XRayMacroBook $sx 'ModesTopLevel' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'A1' = '=T_L1()'
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $bookPath = $book.Path

    $echo = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'TOP'
    if ($echo -notmatch 'DEPTH=TOP') { Complete-Test -Fail -Detail "setter echo: $echo" }

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED \[TOP\]') { Complete-Test -Fail -Detail "arm line: $armLine" }

    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $mark2 = Get-LogLength $paths.Log
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $totLine = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2
    $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

    $problems = @()
    if ($value -ne '7') { $problems += "result wrong: '$value' (expected 7)" }
    if (-not $totLine) { $problems += 'no totals line after disarm' }
    else {
        $t = ConvertFrom-XRayTotals $totLine
        # Counted, not emitted: the accounting must still see all three frames.
        if ($t.framesOpened -lt 3) { $problems += "totals counted only $($t.framesOpened) frames -- ON mode must not thin the accounting" }
        if ($t.framesOpened -ne $t.framesClosed) { $problems += "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
    }
    $vbaEntries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $vbaExits   = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    if ($vbaEntries.Count -ne 1) {
        # Say what the extra rows were: the function names and depths tell a filter that stopped
        # working from rows that belong to another test in a reused session.
        $seen = @($vbaEntries | ForEach-Object { '{0}[depth={1}]' -f $_.function, $_.depth }) -join ','
        $problems += "expected exactly 1 VBA entry row (the top frame), got $($vbaEntries.Count): $seen"
    }
    elseif ($vbaEntries[0].function -ne 'T_L1') { $problems += "top frame is '$($vbaEntries[0].function)', expected T_L1" }
    if ($vbaExits.Count -ne 1) { $problems += "expected exactly 1 VBA exit row, got $($vbaExits.Count)" }
    $inner = @($rows | Where-Object { $_.function -eq 'T_L2' -or $_.function -eq 'T_L3' })
    if ($inner.Count -ne 0) { $problems += "inner frames leaked into a top-level-only trace: $(@($inner | ForEach-Object { $_.function }) -join ',')" }
    $problems += Test-RowInvariants $rows

    if ($problems.Count) { Complete-Test -Fail -Detail ($problems -join '; ') }
    Complete-Test -Pass -Detail "1 of 3 frames emitted, all 3 counted, value=7"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
