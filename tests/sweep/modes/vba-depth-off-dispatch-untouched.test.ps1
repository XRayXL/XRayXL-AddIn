# VBA DEPTH=OFF, the default, leaves the dispatch table untouched and a VBA UDF untraced. The XLL
# side's arm truncates the trace file, so "no VBA rows" is about this arm window.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Public Function T_Quiet(ByVal x As Double) As Double
    T_Quiet = x + 5
End Function
'@
    New-XRayMacroBook $sx 'ModesVbaOff' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'A1' = '=TxB(2,3)'
        'A2' = '=T_Quiet(30)'
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $bookPath = $book.Path

    # Explicit, not just inherited: the default is the assertion here.
    $echo = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF'
    if ($echo -notmatch 'DEPTH=OFF') { Complete-Test -Fail -Detail "setter echo: $echo" }

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    if (-not (Wait-LogLine $paths.Log 'VBA tracing: OFF' $mark)) {
        Complete-Test -Fail -Detail 'arm did not report the VBA side OFF'
    }
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of' $mark
    if (-not $armLine) { Complete-Test -Fail -Detail 'XLL side did not arm' }

    Invoke-XRayRecalc $app
    $xllValue = Get-XRayCellText $ws.Cells.Item(1, 1)
    $vbaValue = Get-XRayCellText $ws.Cells.Item(2, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

    $problems = @()
    if ($xllValue -ne '23') { $problems += "TxB result wrong: '$xllValue'" }
    if ($vbaValue -ne '35') { $problems += "T_Quiet result wrong: '$vbaValue'" }
    $vbaRows = @($rows | Where-Object { $_.source -eq 'VBA' })
    if ($vbaRows.Count -ne 0) { $problems += "$($vbaRows.Count) VBA row(s) traced with VBA tracing OFF" }
    $xllEntries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -and $_.function -eq 'TxB' })
    if ($xllEntries.Count -lt 1) { $problems += 'the XLL side traced nothing -- it should stand alone' }
    $problems += Test-RowInvariants $rows

    if ($problems.Count) { Complete-Test -Fail -Detail ($problems -join '; ') }
    Complete-Test -Pass -Detail "TxB traced, T_Quiet untraced ($($rows.Count) rows, 0 vba)"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
