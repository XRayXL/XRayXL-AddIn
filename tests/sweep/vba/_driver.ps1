# The VBA regression driver: identity, the shadow stack, nesting and recursion. Each .test.ps1
# is self-contained: it carries its own case data and hands it here. Each case gets its own
# workbook built from its own Setup and its own arm/disarm, so the totals are attributable to it
# alone.
#
#    Setup    VBA appended to the standard module (may define several procs)
#    Invoke   what the harness runs, as an Application.Run name + args
#    Then     an optional second call in the same arming session; its target must be
#             defined by this case's own Setup, and a failure to run it fails the test
#    Expect   a scriptblock given the parsed totals; returns $null or a reason
#    Why      what the case is really testing, for the failure message
#
# Expectations are arithmetic, because VBA's own semantics give the answer in advance. A `For i
# = 1 To n ... Next i` body of K statements costs K+1 beginning-of-statement opcodes per
# iteration, plus one for the `For` and one for `End Sub`. Cases assert ranges where the exact
# constant is not the point, and exact equality where it is.

function Invoke-VbaCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')
    $c = $Case

    try {
        $sx = Connect-TestExcel
        $app = $sx.App
        Set-XRaySessionDefaults $sx
        $paths = Get-XRayPaths $sx.ProcId

        # The case table asserts names against the leaf 'VbaRun.xlsm', so the leaf
        # stays fixed and a pid-keyed subdirectory keeps parallel sessions apart.
        # A class or form module goes in before 'Cases', which news them up: the
        # same signature has a different frame in each kind of container.
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

        # Invoke. A cell formula where the case says so -- an unhandled error
        # in a UDF becomes #VALUE! and unwinds silently, where the same error
        # under Application.Run raises the modal VBA dialog. Otherwise
        # Application.Run, QUALIFIED with the workbook, because an unqualified
        # name resolves against whatever Excel considers active.
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

        # An optional SECOND call inside the same arming session: drift only
        # shows up after the disturbance, and a fresh arm resets the shadow
        # stack, so a follow-up in a separate case would test nothing.
        # A follow-up that did not run fails the test, or the case asserts an undisturbed session.
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

        # A case may RAISE on purpose; what still has to hold is the Expect.
        if (-not $ran -and -not $c.Invoke.MayRaise) {
            Complete-Test -Fail -Detail "VBA call failed: $err"
        }
        if (-not $totLine) { Complete-Test -Fail -Detail 'no totals line after disarm' }

        $t = ConvertFrom-XRayTotals $totLine

        # THE TRACER ITSELF MUST NOT HAVE FAULTED. `faults` is a guarded read
        # declining (normal); `hookFaults` is an exception escaping into the
        # hook's own SEH frame, and the breaker opening means it happened
        # repeatedly and tracing stood down.
        if ($t.hookFaults -gt 0) {
            Complete-Test -Fail -Detail "$($t.hookFaults) hook fault(s) -- the tracer faulted inside a VBA thread"
        }
        if ($totLine -match 'BREAKER=OPEN') {
            Complete-Test -Fail -Detail 'circuit breaker OPEN -- tracing stood down'
        }

        $t | Add-Member -NotePropertyName names -NotePropertyValue $names -Force
        $t | Add-Member -NotePropertyName rows -NotePropertyValue $rows -Force

        # Who called the first frame. The driver knows how it invoked the case: a Formula case
        # is called by its cell, an Application.Run case by nothing on a sheet (kind 'none').
        $callerProbs = @(Test-RowInvariants $rows)
        $firstEntry = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') } | Select-Object -First 1)
        # Every case runs a procedure, so a trace with no VBA entry lost it, however the totals read.
        if ($firstEntry.Count -eq 0) { Complete-Test -Fail -Detail "no VBA entry row in this workbook's trace (statements=$($t.statements))" }
        # Calls, when the case knows them: every VBA call in this workbook, in order, and no others.
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
        Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
    }
}
