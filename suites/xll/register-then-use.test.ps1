# REGISTER AN XLL, THEN USE IT IMMEDIATELY -- the call must be traced.
#
# The realistic shape of this is a macro or an add-in loader that calls
# Application.RegisterXLL and, the moment it returns, Application.Run on one of
# the functions that XLL just registered. There is no pause, and no calculation
# in between.
#
# WHAT THIS DOES AND DOES NOT PROVE.
#
# It proves the scenario works: register an XLL, use it at once, get trace rows.
# It does NOT prove there is no race. Registrations are hooked by a worker
# thread, so a window exists in principle -- and it was real: with a 200ms poll
# and a 60ms patch, a use 4ms after RegisterXLL missed every call, 6 times out
# of 6. The poll is now 25ms and the patch 2.4ms, and
# since then no realistic caller has been able to lose the race: PowerShell COM
# marshalling alone takes longer than the window.
#
# So a PASS here means "the window did not open in this run", not "there is no
# window". Written down because a green test on a race is the most misleading
# green there is, and the next person should not read it as a guarantee.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
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

    # The freeze method decides whether patching a registration is affordable at
    # all: the system-wide thread snapshot costs ~60ms per patch, the
    # process-scoped one ~2.4ms. Asserted here because this test's whole subject
    # is how quickly a registration can be hooked.
    #
    # SEARCHED FROM THE TOP OF THE LOG, NOT FROM THIS TEST'S MARK. MinHook is
    # initialised ONCE per PROCESS -- g_minhookReady in src/xll/xllhook.cpp -- so
    # the line is written by whichever arm came first, which under -SessionMode
    # Reuse is an earlier test's. Reading after the mark made a correctly armed
    # session report the slow freeze, and the claim being made here is about
    # the process, not about this arm.
    $freezeLine = Wait-LogLine $paths.Log 'minhook: ' 0 15
    $fastOk = [bool]($freezeLine -match 'process-scoped')
    Write-TestCase -Name 'minhook-process-scoped-freeze' -Pass:$fastOk -Fail:(-not $fastOk) -Detail "$freezeLine"

    # A COPY, because Excel keys a loaded add-in by path, so the same bytes
    # under a new name are a genuinely new module registering afresh.
    $source = (Resolve-Path (Join-Path $PSScriptRoot '..\..\build\x64\Release\TracedAddin\TracedAddin64.xll')).Path
    $second = Join-Path $sx.WorkDir ("UseNow_{0}.xll" -f $sx.ProcId)
    Copy-Item $source $second -Force
    $leaf = Split-Path $second -Leaf

    # ---- THE POINT OF THE TEST: no pause between registering and using ----
    #
    # TWO PATHS, because they race differently. Application.Run carries enough
    # COM overhead that a worker polling every 25ms usually wins -- it passed
    # 3 of 3 while the race was still open, which is the most dangerous kind of
    # green. A cell formula plus CalculateFull is far quicker off the mark and
    # is the one that actually loses: measured 0 of 6 traced. Both are asserted,
    # so passing means the window is closed rather than merely narrow.
    # ORDER MATTERS. The recalc goes FIRST: an Application.Run before it would
    # hand the worker a COM round-trip of head start, and the formula would then
    # be traced for the wrong reason. That mistake was made once here already --
    # the test passed while the race was still open.
    $ws = $app.ActiveSheet
    $ws.Range('A1').Formula = '=TxB(4,5)'      # set BEFORE the XLL loads, so the
    $app.CalculateFull()                       # recalc below is the first thing
    $regOk = [bool]$app.RegisterXLL($second)
    $app.Range('A1').Formula = '=TxB(4,5)'
    $app.CalculateFull()
    $cellValue = Get-XRayCellText $ws.Range('A1')
    $value = $null
    try { $value = $app.Run('TxB', 2, 3) } catch { $value = "EXCEPTION: $($_.Exception.Message)" }
    # -----------------------------------------------------------------------

    Write-TestCase -Name 'second-xll-loaded' -Pass:$regOk -Fail:(-not $regOk) -Detail $second
    if (-not $regOk) { Complete-Test -Fail -Detail "RegisterXLL failed: $second" }

    $valueOk = ("$value" -eq '23')
    Write-TestCase -Name 'call-returned-the-right-answer' -Pass:$valueOk -Fail:(-not $valueOk) `
        -Detail "TxB(2,3) returned '$value', expected 23"

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    # Excel binds the name to the NEWEST registration, so this call goes to the
    # copy. Rows naming the original are counted too, so a failure says which of
    # the two things went wrong: untraced, or traced against the old module.
    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $fromNew = @($rows | Where-Object { $_.module -ieq $leaf -and $_.proc -ieq 'TxB' }).Count
    $fromOld = @($rows | Where-Object { $_.module -ieq 'TracedAddin64.xll' -and $_.proc -ieq 'TxB' }).Count

    # THREE OUTCOMES, AND ONLY ONE OF THEM IS THIS TEST FAILING.
    #
    #   new > 0              the just-registered copy was called AND traced.
    #                        The intended path, and a pass.
    #   new = 0, old > 0     the call went to the ORIGINAL registration. Excel
    #                        had not re-bound the name yet -- both XLLs export
    #                        `TxB`, and which one a call reaches is Excel's
    #                        binding, not our tracing. Four rows came back, so
    #                        nothing was missed; the test simply never got to
    #                        observe its subject. INCONCLUSIVE, not a failure.
    #   new = 0, old = 0     the call was NOT TRACED AT ALL. This is the race
    #                        the file exists for -- the 4ms-after-RegisterXLL
    #                        window that once missed every call, 6 times of 6 --
    #                        and it is a failure.
    #
    # Measured 1 run in 18 landing in the middle case. Failing on it was the
    # test reporting Excel's name binding as a tracer defect, which is the
    # confusion this suite is built to avoid: "we looked and saw nothing" and
    # "there was nothing of ours to see" are different facts.
    $tracedOk    = ($fromNew -gt 0)
    $wentToOld   = (($fromNew -eq 0) -and ($fromOld -gt 0))
    $tracedAtAll = (($fromNew + $fromOld) -gt 0)
    if ($wentToOld) {
        Write-XRayObservation 'immediate-use-is-traced' ("not exercised: Excel called the original registration, not the new copy -- {0} rows from TracedAddin64.xll, 0 from {1}" -f $fromOld, $leaf)
    }
    else {
        Write-TestCase -Name 'immediate-use-is-traced' -Pass:$tracedOk -Fail:(-not $tracedOk) `
            -Detail ("rows from {0}: {1}; rows from the original XLL: {2}" -f $leaf, $fromNew, $fromOld)
    }

    # The cell that was calculated with no pause after RegisterXLL. This is the
    # path that loses the race when registrations are hooked asynchronously.
    $cellRows = @($rows | Where-Object { (Get-CallerCell $_) -ieq 'A1' -and $_.proc -ieq 'TxB' }).Count
    $cellOk = ($cellValue -eq '45') -and ($cellRows -gt 0)
    Write-TestCase -Name 'immediate-recalc-is-traced' -Pass:$cellOk -Fail:(-not $cellOk) `
        -Detail ("A1 showed '$cellValue' (expect 45), trace rows for A1: $cellRows")

    # `$wentToOld` is a SKIP: every other check still has to hold, but the
    # case this file exists for never ran, and a pass would claim it had.
    $runOk = $tracedOk -or $wentToOld
    if ($fastOk -and $valueOk -and $runOk -and $cellOk) {
        if ($wentToOld) {
            Complete-Test -Skip -Detail "Excel called the original registration ($fromOld rows), so immediate use of the new one was not exercised"
        }
        Complete-Test -Pass -Detail "run=$fromNew rows, A1=$cellRows rows"
    }
    else {
        Complete-Test -Fail -Detail "fastFreeze=$fastOk value=$valueOk run-traced=$tracedOk tracedAtAll=$tracedAtAll cell-traced=$cellOk (new=$fromNew old=$fromOld cellRows=$cellRows)"
    }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
finally {
    # The XLL stays loaded, so the delete usually fails; that must never fail
    # the test.
    if ($second) { Remove-Item $second -Force -ErrorAction SilentlyContinue }
}
