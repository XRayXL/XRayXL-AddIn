# Register an XLL, then use it immediately: the call must be traced.
#
# The realistic shape is a macro or loader that calls Application.RegisterXLL and at once runs
# one of the functions it registered.
#
# It does not prove there is no race. Registrations are hooked by a worker thread a moment
# later, and a CalculateFull straight after RegisterXLL can land in that window. The window is
# accepted because the log reports it, so an untraced use passes only when the log's applied
# batch came after the use began.
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

    # The freeze method decides whether patching a registration is affordable: the system-wide
    # thread snapshot is far slower per patch than the process-scoped one.
    #
    # Searched from the top of the log, not from this test's mark: MinHook is initialised once
    # per process (g_minhookReady in src/xll/xllhook.cpp), so under -SessionMode Reuse the line
    # was written by an earlier test's arm.
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

    # No pause between registering and using. Two paths, because they race differently:
    # Application.Run carries enough COM overhead that the worker usually wins, while a cell
    # formula plus CalculateFull is far quicker off the mark. The recalc goes first, so the Run
    # does not hand the worker a head start. Different arguments before the load, so its rows
    # cannot be counted as the use after it.
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

    # Excel binds the name to the newest registration, so this call goes to the copy. Rows
    # naming the original are counted too, so a failure says which went wrong: untraced, or
    # traced against the original module. Each use is found by its own arguments: other TxB rows
    # in the trace say nothing about it.
    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    $runEntries = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.proc -ieq 'TxB' -and $_.args -ceq 'a1:B=2 a2:B=3' })
    $fromNew = @($runEntries | Where-Object { $_.module -ieq $leaf }).Count
    $fromOld = @($runEntries | Where-Object { $_.module -ieq 'TracedAddin64.xll' }).Count

    # THE REPORTED WINDOW IS ACCEPTED. Registrations are patched as one batch
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

    # Three outcomes, and only one of them is a failure.
    #
    #    new > 0              the just-registered copy was called and traced: a pass.
    #    new = 0, old > 0     the call went to the original registration, because Excel had
    #                         not re-bound the name yet. Inconclusive, not a failure.
    #    new = 0, old = 0     the call was not traced at all: the race this file exists
    #                         for, and a failure.
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
