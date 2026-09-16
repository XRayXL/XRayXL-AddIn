# AN ADD-IN UNLOADED AND LOADED AGAIN WHILE ARMED.
#
# XRayXL takes no reference on an add-in it hooks, so unloading one really
# unloads it: holding it would change when someone else's DLL goes away. Loaded
# again at the same address, its code carries no detour while the tracer still
# counts the export as hooked, so the late arm has to notice and patch it again.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$copy = $null
try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    # A copy, so it is a module of its own that can be unloaded without touching the suite's add-in.
    $source = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll')).Path
    # A per-run name: a repeat in a reused session would otherwise meet its own still-loaded copy.
    $copy = Join-Path $sx.WorkDir ("ReloadXll_{0}_{1}.xll" -f $sx.ProcId, [guid]::NewGuid().ToString('N').Substring(0, 8))
    Copy-Item $source $copy -Force
    $leaf = Split-Path $copy -Leaf

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }

    $mark1 = Get-LogLength $paths.Log
    if (-not [bool]$app.RegisterXLL($copy)) { Complete-Test -Fail -Detail "RegisterXLL failed: $copy" }
    $late1 = Wait-LogLine $paths.Log 'late arm:' $mark1 20
    Write-XRayObservation 'first-load' "$late1"

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxB(2,3)'
    Invoke-XRayRecalc $app

    $unregistered = $app.ExecuteExcel4Macro('UNREGISTER("' + $copy + '")')
    $stillLoaded = @([Diagnostics.Process]::GetProcessById($sx.ProcId).Modules |
                     Where-Object { $_.FileName -ieq $copy }).Count -gt 0

    $mark2 = Get-LogLength $paths.Log
    if (-not [bool]$app.RegisterXLL($copy)) { Complete-Test -Fail -Detail "second RegisterXLL failed: $copy" }
    $late2 = Wait-LogLine $paths.Log 'late arm:' $mark2 20
    Write-XRayObservation 'second-load' "$late2"

    $ws.Cells.Item(1, 2).Formula = '=TxB(4,5)'
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 2)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $reloaded = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.module -ieq $leaf -and $_.args -like 'a1:B=4 *' })

    Check 'unloading-really-unloads-it' (-not $stillLoaded) `
          "UNREGISTER returned '$unregistered'; module still loaded: $stillLoaded"
    Check 'the-reloaded-add-in-answered' ($value -eq '45') "B1='$value'"
    Check 'the-reloaded-add-in-is-still-traced' ($reloaded.Count -ge 1) `
          "entries of TxB(4,5) naming $leaf : $($reloaded.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'unloaded for real, loaded again, and traced'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
finally {
    if ($copy) { Remove-Item $copy -Force -ErrorAction SilentlyContinue }
}
