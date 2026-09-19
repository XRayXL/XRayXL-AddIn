# The arm/disarm hammer -- an instrument, not a test.
#
# It hunts an access violation in which execution transfers to an address in no mapped module,
# just after arming, on the first armed calculation. The combination it reproduces:
#
#   * VBA DEPTH=ALL -- the full patch set, every dispatch slot.
#   * Rapid arm/disarm cycles in a reused process.
#   * A volatile UDF, so every calc re-runs it, plus a Worksheet_Calculate
#     handler -- VBA running re-entrantly from inside the calc engine.
#   * A second workbook of VBA open throughout, returning arrays, Variants,
#     objects and strings, so the armed CalculateFull pours a large
#     multi-project workload through the freshly patched slots and runs the
#     return-value decoder on every frame.
#
# To make a rare event happen in minutes: thousands of cycles, multithreaded
# calc so worker threads traverse the patched slots concurrently, both sources
# armed at once, and the capture settings re-randomised every cycle.
#
# The evidence it hunts is written by crashlog's first-chance vectored
# handler: the faulting stack and a minidump.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\sweep\_xray_common.ps1')
. (Join-Path $PSScriptRoot '..\_instrument.ps1')

# The run stops on whichever limit comes first and reports which, because
# cycle cost varies by orders of magnitude between configurations.
#   XRAY_HAMMER_ITERATIONS      the cycle count
#   XRAY_HAMMER_BUDGET_SECONDS  the duration
#   XRAY_HAMMER_CONTROL=1       the same cycles with no arm or disarm: the baseline the resource
#                               samples are read against, since Excel's own growth is in them too
$Iterations = Get-EnvInt 'XRAY_HAMMER_ITERATIONS' 1500
$Control = ($env:XRAY_HAMMER_CONTROL -eq '1')
# Well inside the suite's TestTimeoutSeconds (3000), so the harness never has
# to kill the run and a result is always reported.
$DefaultBudgetSeconds = 2400
# Cycles between resource samples: frequent enough for a trend, rare enough to cost nothing.
$SampleEveryArms = 100
# Excel leaks a pagefile-backed composition surface on some recalculations, armed or not, and
# private bytes do not show it: left alone, a long run exhausts the machine's commit and takes
# other processes with it. Below this much free commit the run stops, and says so.
$MinFreeCommitGB = 4

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # XRAYXL_OUTPUT_DIR moves the output root; the minidump switch is set in suite.psd1
    $logDir = Join-Path (Get-XRayRoot) 'Logs'
    $dumpPath = Join-Path $logDir 'XRayXL_exec_fault.dmp'
    Remove-Item $dumpPath -ErrorAction SilentlyContinue

    # ---- Ballast: a second book of VBA, open for the whole run -------------
    # Every armed CalculateFull recalculates it too, so the exit hook decodes
    # SAFEARRAYs, BSTRs, VARIANTs and interface pointers on the hot path.
    $ballastCode = @'
Public Function B_Arr() As Variant
    B_Arr = Array(1234.5, "two", CLng(3), True)
End Function
Public Function B_Nested() As Variant
    B_Nested = Array(Array(1.5, 2), 3)
End Function
Public Function B_Obj() As Variant
    Set B_Obj = New Collection
End Function
Public Function B_Str() As String
    B_Str = "ballast-" & CStr(Int(Rnd() * 1000))
End Function
Public Function B_Deep(ByVal n As Long) As Double
    If n > 0 Then B_Deep = B_Deep(n - 1) + 1 Else B_Deep = 0
End Function
Public Function B_Vol() As Double
    Application.Volatile
    B_Vol = Rnd()
End Function
'@
    $bcells = @{
        'A1'='=B_Vol()'; 'A2'='=B_Str()'; 'A3'='=B_Deep(9)'; 'A4'='=B_Obj()'
        'C1'='=B_Arr()'; 'F1'='=B_Nested()'
        'A6'='=B_Vol()'; 'A7'='=B_Deep(5)'; 'A8'='=B_Str()'
    }
    New-XRayMacroBook $sx 'HammerBallast' -SheetName 'B1' -Cells $bcells -Components @(
        @{ Kind = 1; Name = 'BM'; Code = $ballastCode }
    )

    # ---- The hammer book: volatile UDFs, a calc event, and the loop --------
    $hammerCode = @'
Public Function H_Vol() As Double
    Application.Volatile
    H_Vol = Rnd()
End Function
Public Function H_Deep(ByVal n As Long) As Double
    If n > 0 Then H_Deep = H_Deep(n - 1) + 1 Else H_Deep = 0
End Function
Public Function H_Chain() As Double
    H_Chain = H_Deep(12) + H_Vol()
End Function
Public Function H_Arr() As Variant
    H_Arr = Array(9.5, "hx", CLng(7), False)
End Function
Public Function H_Obj() As Variant
    Set H_Obj = New Collection
End Function
Public Function H_Str() As String
    H_Str = "h-" & CStr(Int(Rnd() * 1000))
End Function
'@
    $sheetCode = @'
Private Sub Worksheet_Calculate()
    Dim z As Double
    z = H_Deep(6)
End Sub
'@
    # Arrays spill, so each gets its own column.
    $cells = @{
        'A1'='=H_Vol()';  'A2'='=H_Chain()'; 'A3'='=H_Str()';  'A4'='=H_Obj()'
        'A5'='=H_Vol()';  'A6'='=H_Deep(8)'; 'A7'='=H_Vol()';  'A8'='=H_Chain()'
        'C1'='=H_Arr()';  'F1'='=H_Arr()'
        # XLL work in the same book: a plain call, and TxCallsBack, which
        # re-enters Excel through xlUDF so XLL frames nest while the VBA
        # dispatch patches are live.
        'A10'='=TxB(2,3)'; 'A11'='=TxCallsBack(5)'; 'A12'='=TxCallsBack2(5)'
    }
    New-XRayMacroBook $sx 'Hammer' -Cells $cells -Components @(
        @{ Kind = 1; Name = 'M'; Code = $hammerCode }
        @{ Kind = 'Sheet'; Code = $sheetCode }
    )

    # worker threads traverse the shared patched slots concurrently
    $threads = Enable-MultiThreadedCalc $app

    $bad = Set-XRayDepthAll $sx
    if ($bad) { Complete-Test -Fail -Detail $bad }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $died = ''
    $done = 0
    # A shorter budget is a weaker hunt, not a different one. Keep it under the
    # suite's TestTimeoutSeconds: a run the harness kills reports nothing.
    $BudgetSeconds = Get-EnvInt 'XRAY_HAMMER_BUDGET_SECONDS' $DefaultBudgetSeconds
    # Sampled, not asserted: each arm leaks its stub page on purpose (a thread
    # may still be inside it), and there is no honest threshold without a baseline.
    $memFirst = $null; $memLast = $null; $memAt = 0
    # set on every exit path, so a value that cannot be mistaken for an answer is the default
    $stoppedOn = 'unset'

    # The loop is driven from PowerShell over COM: arm, calc and disarm are separate calls.
    try {
        # The decode gates are latched at arm and read on both hot paths, so
        # varying them per cycle varies which code the traced thread runs.
        # Seeded and reported, so a crash is replayable.
        $seed = if ($env:XRAY_HAMMER_SEED) { [int]$env:XRAY_HAMMER_SEED } else { Get-Random -Minimum 1 -Maximum 999999 }
        $rand = [Random]::new($seed)
        $depths = @('TOP', 'ALL')
        for ($i = 1; $i -le $Iterations; $i++) {
            # settings refuse while armed, so they are chosen between disarm and arm
            [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH'  $depths[$rand.Next(0, 2)])
            [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'   ($rand.Next(0, 2) -eq 1))
            [void](Set-XRayTraceParam $sx 'VBA' 'RETVAL' ($rand.Next(0, 2) -eq 1))
            [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH'  $depths[$rand.Next(0, 2)])
            [void](Set-XRayTraceParam $sx 'XLL' 'ARGS'   ($rand.Next(0, 2) -eq 1))
            [void](Set-XRayTraceParam $sx 'XLL' 'RETVAL' ($rand.Next(0, 2) -eq 1))
            if (-not $Control) {
                $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
                if ($pressed -ne 'pressed') { throw "arm refused at $i : $pressed" }
            }
            $app.CalculateFull()
            if (-not $Control) { [void](Invoke-XRayCommand $sx 'XRayXL_Disarm') }
            $done = $i
            if (Test-Path $dumpPath) { $stoppedOn = 'minidump written'; break }   # the fault we are hunting
            if (($i % $SampleEveryArms) -eq 0) {
                $sample = Get-ProcSample -ProcessId $sx.ProcId -At $i
                if ($sample) {
                    $memLast = $sample
                    $memAt = $i
                    if (-not $memFirst) { $memFirst = $memLast }
                    Write-Output ("arms={0,5} working={1}MB private={2}MB handles={3} gdi={4} user={5}" -f `
                        $i, $memLast.WorkingMB, $memLast.PrivateMB, $memLast.Handles, $memLast.Gdi, $memLast.User)
                }
                $freeCommitGB = (Get-CimInstance Win32_OperatingSystem).FreeVirtualMemory / 1MB
                if ($freeCommitGB -lt $MinFreeCommitGB) {
                    $stoppedOn = ('system commit low ({0:N1} GB free)' -f $freeCommitGB)
                    break
                }
            }
            if ($sw.Elapsed.TotalSeconds -ge $BudgetSeconds) { $stoppedOn = 'time budget'; break }
        }
    }
    catch {
        $died = ($_.Exception.Message -replace '\s+', ' ')
        $stoppedOn = 'excel died'
    }
    $sw.Stop()
    if ($stoppedOn -eq 'unset') { $stoppedOn = 'iteration count' }

    # ---- What happened ----------------------------------------------------
    # The action log records every arm, so it counts the cycles survived even
    # when Excel is gone.
    $armed = 0
    try {
        $tail = Get-Content $paths.Log -ErrorAction SilentlyContinue
        $armed = @($tail | Select-String -SimpleMatch 'VBA tracing: ARMED').Count
    } catch {}

    # This Excel's own crash file, by pid (crashlog names it XRayXL_crash_<pid>.txt):
    # a bare glob would pick up another session's crash file.
    $crashFile = @(Get-ChildItem (Join-Path $logDir ("XRayXL_crash_{0}.txt" -f $sx.ProcId)) -ErrorAction SilentlyContinue |
                   Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    $execFault = ''
    if ($crashFile.Count -gt 0) {
        $txt = Get-Content $crashFile[0].FullName -Raw -ErrorAction SilentlyContinue
        if ($txt -match 'EXECUTION IN UNMAPPED MEMORY') {
            $execFault = $crashFile[0].FullName
        }
    }
    $haveDump = Test-Path $dumpPath

    # the slope runs from the first sample, not from zero, so it skips the start-up transient
    $perK = if ($memFirst -and $memAt -gt $memFirst.At) {
        [math]::Round((($memLast.PrivateMB - $memFirst.PrivateMB) / ($memAt - $memFirst.At)) * 1000, 1)
    } else { 0 }
    $memNote = if ($memFirst) {
        ("private {0}->{1}MB over {2} arms ({3}MB per 1000 arms), handles {4}->{5}" -f `
            $memFirst.PrivateMB, $memLast.PrivateMB, $memAt, $perK, $memFirst.Handles, $memLast.Handles)
    } else { "no samples" }
    $detail = ("{7}iterations={0} completed={1} armed={2} secs={3} mtcThreads={4} seed={5} stoppedOn={6} | $memNote" -f
               $Iterations, $done, $armed, [math]::Round($sw.Elapsed.TotalSeconds, 1), $threads, $seed, $stoppedOn,
               $(if ($Control) { 'CONTROL (never armed) ' } else { '' }))
    if ($died) { $detail += "; excel died: $died" }
    if ($execFault) { $detail += "; EXEC-FAULT CAPTURED: $execFault" }
    if ($haveDump) { $detail += "; dump: $dumpPath" }

    # a crash is the finding this instrument exists for, so it fails, with where the evidence is
    if ($died -or $execFault) { Complete-Test -Fail -Detail $detail }
    Complete-Test -Pass -Detail ("survived -- " + $detail)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
