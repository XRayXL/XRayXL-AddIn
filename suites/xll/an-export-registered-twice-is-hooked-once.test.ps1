# ONE EXPORT, TWO REGISTRATIONS.
#
# TracedAddin registers TxB a second time, as TxBAgain. Both names reach the same
# code, so arming hooks it once; the second registration is not a detour that
# failed.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }
    $detour = if ($armLine -match 'detour (\d+)') { [int]$Matches[1] } else { -1 }

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxBAgain(2,3)'
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) |
              Where-Object { $_.kind -eq 'entry' -and $_.args -eq 'a1:B=2 a2:B=3' })

    Check 'no-detour-counted-as-failed' ($detour -eq 0) "$armLine"
    Check 'the-second-name-answered' ($value -eq '23') "A1='$value'"
    Check 'a-call-through-the-second-name-is-traced' ($rows.Count -ge 1) "entries with a1:B=2 a2:B=3: $($rows.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "hooked once, traced through both names"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
