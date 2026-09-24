# The VBA case driver. Each case gets its own workbook and its own arm/disarm, so the totals are
# attributable to it alone.
#
#    Setup    VBA appended to the standard module (may define several procs)
#    Invoke   what the harness runs, as an Application.Run name + args
#    Then     an optional second call in the same arming session; its target must be
#             defined by this case's own Setup, and a failure to run it fails the test
#    Expect   a scriptblock given the parsed totals; returns $null or a reason
#    Why      what the case is really testing, for the failure message
#
# Expectations are arithmetic, because VBA's own semantics give the answer in advance.

function Invoke-VbaCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')
    $c = $Case

    try {
        $sx = Connect-TestExcel
        $app = $sx.App
        Set-XRaySessionDefaults $sx
        $paths = Get-XRayPaths $sx.ProcId

        # Cases assert names against the fixed leaf 'VbaRun.xlsm'. Class and form go in before
        # 'Cases', which news them up.
        $components = @()
        if ($c.ClassSetup) { $components += @{ Kind = 2; Name = 'CCase'; Code = $c.ClassSetup } }
        if ($c.FormSetup)  { $components += @{ Kind = 3; Name = 'UFCase'; Code = $c.FormSetup } }
        $components += @{ Kind = 1; Name = 'Cases'; Code = $c.Setup }
        New-XRayMacroBook $sx 'vba' $components -SheetName 'Sheet1' -Leaf 'VbaRun.xlsm'
        $book = Get-XRayMacroBook
        $ws = $book.Sheet; $bookPath = $book.Path; $bookLeaf = $book.Leaf

        $mark = Get-LogLength $paths.Log
        $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
        if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
        $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
        if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

        # A formula case exists because an unhandled error in a UDF unwinds silently, where under
        # Application.Run it raises a modal dialog. Run is qualified: a bare name resolves against
        # whatever Excel considers active.
        $ran = $true; $err = ''
        try {
            if ($c.Invoke.Formula) {
                $ws.Range('H1').Formula = $c.Invoke.Formula
                Invoke-XRayRecalc $app
            }
            else {
                $qualified = "$bookLeaf!" + $c.Invoke.Name
                if ($c.Invoke.Args.Count -eq 0) { $app.Run($qualified) | Out-Null }
                else { $app.Run($qualified, $c.Invoke.Args[0]) | Out-Null }
            }
        } catch { $ran = $false; $err = $_.Exception.Message }

        # Same arming session: drift shows only after the disturbance, and a fresh arm resets the
        # shadow stack. A follow-up that did not run fails, or the case would assert nothing.
        if ($c.Then) {
            try {
                $q2 = "$bookLeaf!" + $c.Then.Name
                if (-not $c.Then.Args -or $c.Then.Args.Count -eq 0) { $app.Run($q2) | Out-Null }
                else { $app.Run($q2, $c.Then.Args[0]) | Out-Null }
            } catch {
                Complete-Test -Fail -Detail ("follow-up call $($c.Then.Name) failed: " +
                                             $_.Exception.Message +
                                             ' -- Then must name a procedure this case defines')
            }
        }

        $mark2 = Get-LogLength $paths.Log
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        $totLine = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2
        $names = Read-XRayNames $paths.Log $mark2
        $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

        # A case may raise on purpose; what still has to hold is the Expect.
        if (-not $ran -and -not $c.Invoke.MayRaise) {
            Complete-Test -Fail -Detail "VBA call failed: $err"
        }
        if (-not $totLine) { Complete-Test -Fail -Detail 'no totals line after disarm' }

        $t = ConvertFrom-XRayTotals $totLine

        # `faults` is a guarded read declining, which is normal; `hookFaults` is an exception
        # reaching the hook's SEH frame, and an open breaker means tracing stood down.
        if ($t.hookFaults -gt 0) {
            Complete-Test -Fail -Detail "$($t.hookFaults) hook fault(s) -- the tracer faulted inside a VBA thread"
        }
        if ($totLine -match 'BREAKER=OPEN') {
            Complete-Test -Fail -Detail 'circuit breaker OPEN -- tracing stood down'
        }

        $t | Add-Member -NotePropertyName names -NotePropertyValue $names -Force
        $t | Add-Member -NotePropertyName rows -NotePropertyValue $rows -Force

        # A formula case is called by its cell, an Application.Run case by nothing on a sheet.
        $callerProbs = @(Test-RowInvariants $rows)
        $firstEntry = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') } | Select-Object -First 1)
        # Every case runs a procedure, so a trace with no VBA entry lost it, however the totals read.
        if ($firstEntry.Count -eq 0) { Complete-Test -Fail -Detail "no VBA entry row in this workbook's trace (statements=$($t.statements))" }
        if ($c.ContainsKey('Calls')) { $callerProbs += Test-ExpectedTrace $rows $c.Calls 'VBA' }
        if ($firstEntry.Count -eq 1) {
            $expectCaller = if ($c.Invoke.Formula) { 'cell' } else { 'none' }
            if ($firstEntry[0].caller -ne $expectCaller) {
                $callerProbs += "first frame's caller is '$($firstEntry[0].caller)', expected '$expectCaller'"
            }
        }
        if ($callerProbs.Count) { Complete-Test -Fail -Detail ($callerProbs -join '; ') }

        $why = & $c.Expect $t
        if ($why) { Complete-Test -Fail -Detail $why }

        $raised = if (-not $ran) { ' raised-as-intended' } else { '' }
        Complete-Test -Pass -Detail ("stmts={0} procs={1} depth={2} frames={3}/{4} exits={5} faults={6}{7}" -f `
            $t.statements, $t.procedures, $t.maxDepth, $t.framesOpened, $t.framesClosed, $t.exits, $t.faults, $raised)
    }
    catch {
        Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
    }
}
