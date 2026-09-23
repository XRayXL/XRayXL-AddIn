# XLL DEPTH=OFF means NOT HOOKED: arming touches nothing on
# the XLL side, the formula still calculates correctly, and the trace holds
# no XLL rows -- while the VBA side, set to DEPTH=ALL, arms and traces on its
# own, exactly as a VBA-only workbook would. The VBA side also proves the
# trace file is THIS session's (its arm truncates it), so "no XLL rows" is
# an assertion about this arm window, not about a stale file.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Public Function T_Plain(ByVal x As Double) As Double
    T_Plain = x * 2
End Function
'@
    New-XRayMacroBook $sx 'ModesXllOff' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{
        'A1' = '=TxB(2,3)'
        'A2' = '=T_Plain(21)'
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $bookPath = $book.Path

    $echo = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF'
    if ($echo -notmatch 'DEPTH=OFF') { Complete-Test -Fail -Detail "xll setter echo: $echo" }
    $echo = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'ALL'
    if ($echo -notmatch 'DEPTH=ALL') { Complete-Test -Fail -Detail "vba setter echo: $echo" }

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    if (-not (Wait-LogLine $paths.Log 'XLL tracing: OFF' $mark)) {
        Complete-Test -Fail -Detail 'arm did not report the XLL side OFF'
    }
    if (-not (Wait-LogLine $paths.Log 'VBA tracing: ARMED' $mark)) {
        Complete-Test -Fail -Detail 'VBA side did not arm alongside the disabled XLL side'
    }

    Invoke-XRayRecalc $app
    $xllValue = Get-XRayCellText $ws.Cells.Item(1, 1)
    $vbaValue = Get-XRayCellText $ws.Cells.Item(2, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

    $problems = @()
    if ($xllValue -ne '23') { $problems += "TxB result wrong: '$xllValue'" }
    if ($vbaValue -ne '42') { $problems += "T_Plain result wrong: '$vbaValue'" }
    $xllRows = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -or ($_.kind -eq 'exit' -and $_.source -eq 'XLL') })
    if ($xllRows.Count -ne 0) { $problems += "$($xllRows.Count) XLL row(s) traced with XLL tracing OFF" }
    $vbaRows = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'T_Plain' })
    if ($vbaRows.Count -lt 1) { $problems += 'the VBA side traced nothing -- it should stand alone' }
    $problems += Test-RowInvariants $rows

    if ($problems.Count) { Complete-Test -Fail -Detail ($problems -join '; ') }
    Complete-Test -Pass -Detail "TxB=23 untraced, T_Plain traced ($($rows.Count) rows, 0 xll)"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
