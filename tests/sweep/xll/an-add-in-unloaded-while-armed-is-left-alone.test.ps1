# AN ADD-IN UNLOADED WHILE ARMED IS LEFT ALONE, AND TRACED AGAIN WHEN IT RETURNS.
#
# Disarm puts each hooked function's first bytes back. For an add-in unloaded
# while armed there is nothing of its to restore, or another module now sits at
# that address and must not be written to, so its detours are left untouched. When
# the same add-in loads again later, arming patches it afresh.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$copy = $null
try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    $source = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll')).Path
    # A per-run name: a repeat in a reused session would otherwise meet its own still-loaded copy.
    $copy = Join-Path $sx.WorkDir ("UnloadXll_{0}_{1}.xll" -f $sx.ProcId, [guid]::NewGuid().ToString('N').Substring(0, 8))
    Copy-Item $source $copy -Force
    $leaf = Split-Path $copy -Leaf
    $ws = $app.ActiveSheet

    # ---- hooked late, then unloaded while armed ------------------------------
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    [void](Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark)
    $mark1 = Get-LogLength $paths.Log
    if (-not [bool]$app.RegisterXLL($copy)) { Complete-Test -Fail -Detail "RegisterXLL failed: $copy" }
    [void](Wait-LogLine $paths.Log 'late arm:' $mark1 20)
    $ws.Cells.Item(1, 1).Formula = '=TxB(2,3)'
    Invoke-XRayRecalc $app

    [void]$app.ExecuteExcel4Macro('UNREGISTER("' + $copy + '")')
    $stillLoaded = @([Diagnostics.Process]::GetProcessById($sx.ProcId).Modules |
                     Where-Object { $_.FileName -ieq $copy }).Count -gt 0

    $mark2 = Get-LogLength $paths.Log
    $drops = Invoke-XRayDisarm $sx
    [void](Wait-LogLine $paths.Log 'disarmed' $mark2 20)
    $log = [IO.File]::ReadAllText($paths.Log)
    $disarmLog = $log.Substring([Math]::Min($mark2, $log.Length))

    # ---- loaded again while disarmed, then armed -----------------------------
    if (-not [bool]$app.RegisterXLL($copy)) { Complete-Test -Fail -Detail "second RegisterXLL failed: $copy" }
    $ws.Cells.Item(1, 2).Formula = '=TxB(4,5)'
    $mark3 = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armAgain = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark3
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 2)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $returned = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.module -ieq $leaf -and $_.args -like 'a1:B=4 *' })

    Check 'unloading-really-unloads-it' (-not $stillLoaded) "module still loaded after UNREGISTER: $stillLoaded"
    Check 'disarm-left-the-unloaded-add-in-untouched' ($disarmLog -match 'left untouched') `
          "disarm log: $(($disarmLog -split "`n" | Where-Object { $_ -match 'disarm' }) -join ' | ')"
    Check 'disarm-reported-nothing-stuck' (($disarmLog -notmatch 'still hooked') -and ($drops -eq 0)) `
          "XRayXL_Disarm returned $drops"
    Check 'it-arms-again' ($armAgain -match 'armed \d+ of') "$armAgain"
    Check 'the-returned-add-in-answered' ($value -eq '45') "B1='$value'"
    Check 'the-returned-add-in-is-traced' ($returned.Count -ge 1) "entries of TxB(4,5) naming $leaf : $($returned.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'left untouched while gone, traced again when it returned'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
finally {
    if ($copy) { Remove-Item $copy -Force -ErrorAction SilentlyContinue }
}
