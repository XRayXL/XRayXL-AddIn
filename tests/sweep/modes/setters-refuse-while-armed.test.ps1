# The setter refuses while armed: the modes are read once at arm, so a mid-session change must be
# refused loudly, change nothing, and work again once disarmed.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $bookPath = Join-Path $sx.WorkDir ("ModesRefuse_{0}.xlsx" -f $sx.ProcId)
    Close-OwnLeftover $app (Split-Path $bookPath -Leaf)   # a reused session may still hold our previous one
    Remove-Item $bookPath -ErrorAction SilentlyContinue
    $booksRef = $app.Workbooks
    $wb = $booksRef.Add(); try { $wb.EnableAutoRecover = $false } catch {}
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'S1'
    $ws.Range('A1').Formula = '=TxB(2,3)'
    $wb.SaveAs($bookPath, 51); $wb.Close($false)
    foreach ($r in @($ws, $wb)) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($r) } catch {} }
    $wb = $booksRef.Open($bookPath)
    try { $wb.EnableAutoRecover = $false } catch {}

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of' $mark
    if (-not $armLine) { Complete-Test -Fail -Detail 'did not arm' }

    $failed = 0
    $e1 = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF'
    $ok = ($e1 -match '#Err - cannot change settings while armed')
    Write-TestCase 'xll-setter-refused-while-armed' -Pass:$ok -Fail:(-not $ok) -Detail $e1
    if (-not $ok) { $failed++ }

    $e2 = Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'ALL'
    $ok = ($e2 -match '#Err - cannot change settings while armed')
    Write-TestCase 'vba-setter-refused-while-armed' -Pass:$ok -Fail:(-not $ok) -Detail $e2
    if (-not $ok) { $failed++ }

    # Reading is not a change and must still answer while armed.
    $q = [string](Get-XRayTraceParam $sx 'XLL' 'DEPTH')
    $ok = ($q -eq 'ALL')
    Write-TestCase 'query-still-answers-while-armed' -Pass:$ok -Fail:(-not $ok) -Detail $q
    if (-not $ok) { $failed++ }

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $e3 = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF'
    $ok = ($e3 -match 'XLL DEPTH=OFF')
    Write-TestCase 'setter-works-again-after-disarm' -Pass:$ok -Fail:(-not $ok) -Detail $e3
    if (-not $ok) { $failed++ }

    if ($failed) { Complete-Test -Fail -Detail "$failed refusal case(s) failed" }
    Complete-Test -Pass
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
