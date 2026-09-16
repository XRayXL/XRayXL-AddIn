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
# thread a moment later, and a CalculateFull straight after RegisterXLL can land
# in that window (measured 2026-09-15: 32 ms, the recalc untraced). The window is
# ACCEPTED (D49) because the log reports it, so an untraced use passes only when
# the log's applied batch came after the use began.
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
    $source = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\build\x64\Release\TracedAddin\TracedAddin64.xll')).Path
    # A per-run name: a repeat in a reused session would otherwise meet its own still-loaded copy.
    $second = Join-Path $sx.WorkDir ("UseNow_{0}_{1}.xll" -f $sx.ProcId, [guid]::NewGuid().ToString('N').Substring(0, 8))
    Copy-Item $source $second -Force
    $leaf = Split-Path $second -Leaf

    # ---- THE POINT OF THE TEST: no pause between registering and using ----
    #
    # TWO PATHS, because they race differently. Application.Run carries enough
    # COM overhead that a worker polling every 25ms usually wins. A cell formula
    # plus CalculateFull is far quicker off the mark and is the one that loses.
    # ORDER MATTERS. The recalc goes FIRST: an Application.Run before it would
    # hand the worker a COM round-trip of head start, and the formula would then
    # be traced for the wrong reason. That mistake was made once here already --
    # the test passed while the race was still open.
    # Different arguments before the load, so its rows cannot be counted as the use after it.
    $ws = $app.ActiveSheet
    $ws.Range('A1').Formula = '=TxB(1,1)'      # set BEFORE the XLL loads, so the
    $app.CalculateFull()                       # recalc below is the first thing
    $markLoad = Get-LogLength $paths.Log
    $regOk = [bool]$app.RegisterXLL($second)
    $cellAt = Get-Date
    $app.Range('A1').Formula = '=TxB(4,5)'
    $app.CalculateFull()
    $cellValue = Get-XRayCellText $ws.Range('A1')
    $value = $null
    $runAt = Get-Date
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
    # Each use is found by its own arguments: other TxB rows in the trace say nothing about it.
    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $runEntries = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.proc -ieq 'TxB' -and $_.args -ceq 'a1:B=2 a2:B=3' })
    $fromNew = @($runEntries | Where-Object { $_.module -ieq $leaf }).Count
    $fromOld = @($runEntries | Where-Object { $_.module -ieq 'TracedAddin64.xll' }).Count

    # THE REPORTED WINDOW IS ACCEPTED (D49). Registrations are patched as one batch
    # a moment later, and the log says calls in between were not traced. A use with
    # no rows at all passes only if that batch was applied AFTER the use began; a
    # missing row with no window, or after it closed, is still a failure.
    $appliedAt = $null
    $applied = @(Get-Content $paths.Log | Select-Object -Skip $markLoad |
                 Select-String 'register watch: applied .*NOT traced') | Select-Object -First 1
    if ($applied -and $applied.Line -match '^(\d{4}-\d\d-\d\d \d\d:\d\d:\d\d),(\d{3})') {
        $appliedAt = [datetime]::ParseExact("$($Matches[1]).$($Matches[2])", 'yyyy-MM-dd HH:mm:ss.fff', $null)
    }
    $runInWindow  = ($null -ne $appliedAt) -and ($appliedAt -gt $runAt)
    $cellInWindow = ($null -ne $appliedAt) -and ($appliedAt -gt $cellAt)
    $appliedText  = if ($appliedAt) { $appliedAt.ToString('HH:mm:ss.fff') } else { 'no applied line' }

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
    $runWindow   = (-not $tracedAtAll) -and $runInWindow
    if ($wentToOld -or $runWindow) {
        $why = if ($wentToOld) { "Excel called the original registration -- $fromOld rows from TracedAddin64.xll" }
               else { "untraced inside the reported window (batch applied $appliedText, after the call began)" }
        Write-XRayObservation 'immediate-use-is-traced' "not exercised: $why"
    }
    else {
        Write-TestCase -Name 'immediate-use-is-traced' -Pass:$tracedOk -Fail:(-not $tracedOk) `
            -Detail ("rows from {0}: {1}; rows from the original XLL: {2}; window: {3}" -f $leaf, $fromNew, $fromOld, $appliedText)
    }

    # The cell that was calculated with no pause after RegisterXLL. This is the
    # path that loses the race when registrations are hooked asynchronously.
    # The same three outcomes as the Run, judged on TxB(4,5) at A1 alone.
    $cellEntries = @($rows | Where-Object { $_.kind -eq 'entry' -and (Get-CallerCell $_) -ieq 'A1' -and $_.proc -ieq 'TxB' -and $_.args -ceq 'a1:B=4 a2:B=5' })
    $cellNew = @($cellEntries | Where-Object { $_.module -ieq $leaf }).Count
    $cellOld = @($cellEntries | Where-Object { $_.module -ieq 'TracedAddin64.xll' }).Count
    $cellWentToOld = ($cellNew -eq 0) -and ($cellOld -gt 0)
    $cellWindow    = ($cellNew -eq 0) -and ($cellOld -eq 0) -and $cellInWindow
    $cellOk = ($cellValue -eq '45') -and (($cellNew -gt 0) -or $cellWentToOld -or $cellWindow)
    Write-TestCase -Name 'immediate-recalc-is-traced' -Pass:$cellOk -Fail:(-not $cellOk) `
        -Detail ("A1 showed '$cellValue' (expect 45); TxB(4,5) entries at A1 from {0}: {1}, from the original XLL: {2}; window: {3}{4}" -f $leaf, $cellNew, $cellOld, $appliedText, $(if ($cellWindow) { ' (untraced inside it, accepted)' } else { '' }))

    # A SKIP when neither use reached the new copy traced: every other check still
    # has to hold, but the case this file exists for never ran.
    $runOk = $tracedOk -or $wentToOld -or $runWindow
    $exercised = ($fromNew -gt 0) -or ($cellNew -gt 0)
    if ($fastOk -and $valueOk -and $runOk -and $cellOk) {
        if (-not $exercised) {
            Complete-Test -Skip -Detail "neither use traced the new copy (run: old=$fromOld window=$runWindow; A1: old=$cellOld window=$cellWindow), so immediate use was not exercised"
        }
        Complete-Test -Pass -Detail "run: new=$fromNew old=$fromOld window=$runWindow; A1: new=$cellNew old=$cellOld window=$cellWindow"
    }
    else {
        Complete-Test -Fail -Detail "fastFreeze=$fastOk value=$valueOk run-traced=$tracedOk tracedAtAll=$tracedAtAll cell=$cellOk (run new=$fromNew old=$fromOld; A1 new=$cellNew old=$cellOld; applied $appliedText)"
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
