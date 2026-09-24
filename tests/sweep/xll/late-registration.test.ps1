# An XLL loaded after arming must still be traced: arming hooks only what is registered then, so
# later registrations are caught at xlfRegister and hooked from a worker thread. The last check is
# the one that matters: a row naming the new module, since the log line alone would pass an untaken patch.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$second = $null
try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }

    $watchLine = Wait-LogLine $paths.Log 'register watch: (installed|NOT installed|DISABLED)' $mark 15
    $watchOk = [bool]($watchLine -match 'installed on MdCallBack12')
    Write-TestCase -Name 'watch-installed' -Pass:$watchOk -Fail:(-not $watchOk) -Detail "$watchLine"
    if (-not $watchOk) { Complete-Test -Fail -Detail 'the register watch did not install' }

    # A copy, because Excel keys a loaded add-in by path. Named by pid and a per-run suffix, so
    # neither a parallel worker nor an earlier run's still-loaded copy collides.
    $source = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll')).Path
    $second = Join-Path $sx.WorkDir ("LateXll_{0}_{1}.xll" -f $sx.ProcId, [guid]::NewGuid().ToString('N').Substring(0, 8))
    Copy-Item $source $second -Force

    $mark2 = Get-LogLength $paths.Log
    $regOk = [bool]$app.RegisterXLL($second)
    Write-TestCase -Name 'second-xll-loaded' -Pass:$regOk -Fail:(-not $regOk) -Detail $second
    if (-not $regOk) { Complete-Test -Fail -Detail "RegisterXLL failed: $second" }

    $lateLine = Wait-LogLine $paths.Log 'late arm:.*hooked \d+|late arm:.*could not' $mark2 20
    $lateOk = [bool]($lateLine -match 'late arm:.*hooked [1-9]')
    Write-TestCase -Name 'late-registrations-hooked' -Pass:$lateOk -Fail:(-not $lateOk) -Detail "$lateLine"

    # The one that cannot be faked by a log line: call it, and require the trace
    # to name the module that was loaded after arming.
    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxB(2,3)'
    Invoke-XRayRecalc $app
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $leaf = Split-Path $second -Leaf
    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId) | Where-Object { $_.module -ieq $leaf })
    $tracedOk = ($rows.Count -gt 0)
    Write-TestCase -Name 'late-xll-actually-traced' -Pass:$tracedOk -Fail:(-not $tracedOk) `
        -Detail ("rows naming {0}: {1}" -f $leaf, $rows.Count)

    if ($watchOk -and $lateOk -and $tracedOk) { Complete-Test -Pass -Detail "late-xll rows=$($rows.Count)" }
    else { Complete-Test -Fail -Detail "watch=$watchOk late=$lateOk traced=$tracedOk" }
}
catch {
    # Without this an exception ends the script with no verdict, which the manager reports as ERROR.
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
finally {
    # The XLL stays loaded, so the delete usually fails -- that is expected and
    # must never fail the test.
    if ($second) { Remove-Item $second -Force -ErrorAction SilentlyContinue }
}
