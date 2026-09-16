<#
.SYNOPSIS
    StretchXL -- a generic harness for testing things that run inside Excel,
    built around one question: does an Excel session start, do what it is
    asked, and shut down cleanly?

.DESCRIPTION
    See StretchXL.md for the design, the contract and the locked decisions.
    In short: the manager owns every Excel lifecycle -- start, identity,
    ledger, deadlines, measured close, dumps -- and each test is its own file
    in a folder tree, run as its own process, bound to the manager's Excel by
    window handle. With no -Path, the floor runs: start Excel, add a workbook,
    close -- the control that proves the infrastructure clean on its own.

    Results stream to STDOUT one line per result as it happens, and land in a
    results-<timestamp>.jsonl in -OutDir, whose first line records the full
    configuration. Nothing here is written with Write-Host: that goes to the
    information stream, which a pipeline or a scraper never sees.

    Exit code: 0 only if every result was benign. Any FAIL, TIMEOUT, ERROR,
    SETUP-FAILED verdict, any CRASH, HANG, EXITED or FAILED outcome, or fewer
    results than expected, exits 1 -- CI can gate on the exit code with no
    parsing.

.PARAMETER Parallel
    Number of Excel sessions running at once. Mandatory -- concurrency is not
    a detail here. It changes timing, and timing is what this is measuring.

.PARAMETER Runs
    How many times each selected test runs (or, at the floor, the TOTAL number
    of floor iterations). Workers PARTITION the work; they do not multiply it.
    A soak is a high -Runs, usually with -RandomOrder.

.PARAMETER Path
    The tests root: every *.test.ps1 under it (recursively) is a test, and the
    folder is the selection mechanism -- point at a subfolder to run a subset.
    Omitted, the floor runs instead: start Excel, add one empty workbook,
    close. Nothing else at all.

.PARAMETER OutDir
    Where results go. REQUIRED on every run, floor included -- no hidden temp
    locations. Each invocation writes its own results-<timestamp>.jsonl here;
    the first line is a metadata record (full configuration, seed, machine,
    Excel build). Dumps also land here unless -DumpDir says otherwise.

.PARAMETER DumpDir
    Optional separate folder for minidumps. Defaults to -OutDir.

.PARAMETER SessionMode
    Fresh (default): every test gets its own Excel -- isolation, and a per-test
    shutdown measurement.
    Reuse: a sequence of tests shares one long-lived Excel PROCESS. Its
    Application settings, loaded XLLs and in-process VBA / XLL state persist --
    that dirt is what "fails only in a dirty session" is made of -- but its
    WORKBOOKS do not: every test closes the books it opened (Complete-Test),
    so a book is never left for the next test's CalculateFull to recalculate
    into its trace.
    ReuseClean: the same, and ALSO restores the Application settings to
    baseline between tests -- reuse without the settings dirt.
    A session is only shared by tests whose set-up matches: the same
    RegisterXll list, RegisterXllSettleSeconds and SessionEnvironment.
    Reuse and ReuseClean need -Path (the floor run is the fresh-session
    control by definition).

.PARAMETER CloseTimeoutSeconds
    How long to wait for Excel to exit after the close before calling the
    session a HANG. A classifier, not a measurement: every session reports its
    real close-to-exit time in shutdown=. A healthy floor shutdown is under
    three seconds; a deadline shorter than the slowest clean shutdown does not
    measure hangs, it manufactures them.

.PARAMETER TestTimeoutSeconds
    How long a test body may run before it is a TIMEOUT: both the Excel and
    the wedged test process are dumped (the test's stack names the stuck COM
    call, Excel's names why), then the test process is killed. Distinct from
    -CloseTimeoutSeconds because "the test never returned" and "Excel would
    not exit afterwards" are different failures.

.PARAMETER CloseWithX
    Close Excel the way a user does -- make the window visible and post it the
    message the X button posts (WM_SYSCOMMAND / SC_CLOSE) -- instead of
    calling Application.Quit() over COM. The two run different teardown code;
    an add-in that hangs close-down for a user can look healthy under Quit,
    and the reverse. Manager-only by design: running the same suite both ways
    is the point. The window is shown minimized WITHOUT ACTIVATION
    (SW_SHOWMINNOACTIVE), measured not to move focus, so a soak can run while
    you work; the interactive close path still engages.

.PARAMETER Warmup
    Each worker performs ONE unmeasured bare floor close before its counted
    work. Exists because of a measured property of the X-click arm: the first
    X-closed Excel of each client process shuts down in ~60s (Click-to-Run
    implicated); every later one takes ~2.5s. Warm-ups run in parallel, so
    all workers warming at once costs ~60s of wall clock in total. Reported
    as run=warmup lines -- a warm-up that hangs is still loud -- but never
    counted. Never includes suite set-up: set-up costs stay in measured data.

.PARAMETER RandomOrder
    Shuffle the selected work. The shuffle is seeded, the seed is printed in
    the header and recorded in the metadata line, and -Seed replays it -- a
    random order that found a failure and cannot be replayed is half a result.

.PARAMETER Seed
    Replay a specific shuffle. Ignored without -RandomOrder.

.PARAMETER GroupBySession
    Keep tests that share a session set-up (add-ins, settle time, environment)
    together, so a reused session lasts through its whole group instead of
    being replaced whenever the next test needs different add-ins. Order within
    a group is kept; with -RandomOrder the group order is shuffled by the same
    seed. ON BY DEFAULT for -RandomOrder with -SessionMode Reuse or ReuseClean;
    give it to group an unshuffled reuse run. Needs -Path and a reuse mode.

.PARAMETER NoGroupBySession
    Turn the default grouping off: a shuffled reuse run mixes every test into
    every session, as before -GroupBySession existed. Replaying a seed from such
    a run needs this switch again.

.PARAMETER DumpAfterSeconds
    If Excel is still alive this many seconds after the close, write a
    minidump WHILE IT IS STILL STUCK, then carry on waiting. 0 disables it.
    A HANG is always dumped before it is killed; this additionally catches
    slow-but-eventually-clean shutdowns that a dump at the deadline never sees.

        python StretchXL\evidence\hangwhere.py <dump>

.PARAMETER FullDump
    Full-memory dumps instead of stacks-and-handles (about 20x the size).

.PARAMETER NoDump
    Never dump, not even on HANG. For soaks where disk is the constraint.

.PARAMETER Cleanup
    Kill any Excel left behind by an interrupted run, using the pid ledgers,
    and exit. Takes no other action.

.EXAMPLE
    .\StretchXL.ps1 -Parallel 1 -OutDir C:\sx\results
    The floor control run: start Excel, add a workbook, close, once.

.EXAMPLE
    .\StretchXL.ps1 -Parallel 2 -Runs 5 -Path .\suites -OutDir C:\sx\results
    Every test under suites\, five times each, two sessions at a time.

.EXAMPLE
    .\StretchXL.ps1 -Parallel 4 -Runs 50 -Path .\suites\proof -OutDir C:\sx\r -CloseWithX -Warmup -RandomOrder
    A shuffled X-close soak of one suite, warm-ups paid up front.

.EXAMPLE
    .\StretchXL.ps1 -Parallel 8 -Runs 6 -Path .\suites -OutDir C:\sx\r -SessionMode Reuse -RandomOrder
    A shuffled soak whose reused sessions each last through a whole group of tests
    (grouping is the default here; -NoGroupBySession mixes every test instead).

.EXAMPLE
    .\StretchXL.ps1 -Cleanup
    Clear orphans from a run that was interrupted.
#>
[CmdletBinding(DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')][ValidateRange(1, 32)][int]$Parallel,
    [Parameter(ParameterSetName = 'Run')][ValidateRange(1, 100000)][int]$Runs = 1,
    [Parameter(ParameterSetName = 'Run')][string]$Path,
    [Parameter(Mandatory = $true, ParameterSetName = 'Run')][string]$OutDir,
    [Parameter(ParameterSetName = 'Run')][string]$DumpDir,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('Fresh', 'Reuse', 'ReuseClean')][string]$SessionMode = 'Fresh',
    [Parameter(ParameterSetName = 'Run')][ValidateRange(5, 3600)][int]$CloseTimeoutSeconds = 70,
    [Parameter(ParameterSetName = 'Run')][ValidateRange(5, 86400)][int]$TestTimeoutSeconds = 120,
    [Parameter(ParameterSetName = 'Run')][switch]$CloseWithX,
    [Parameter(ParameterSetName = 'Run')][switch]$Warmup,
    [Parameter(ParameterSetName = 'Run')][switch]$RandomOrder,
    [Parameter(ParameterSetName = 'Run')][int]$Seed = -1,
    [Parameter(ParameterSetName = 'Run')][switch]$GroupBySession,
    [Parameter(ParameterSetName = 'Run')][switch]$NoGroupBySession,
    [Parameter(ParameterSetName = 'Run')][ValidateRange(0, 3600)][int]$DumpAfterSeconds = 0,
    [Parameter(ParameterSetName = 'Run')][switch]$FullDump,
    [Parameter(ParameterSetName = 'Run')][switch]$NoDump,
    [Parameter(Mandatory = $true, ParameterSetName = 'Cleanup')][switch]$Cleanup
)

$ErrorActionPreference = 'Stop'

# Shared with TestKit.ps1; each job worker dot-sources the same file by this path.
$commonScript = Join-Path $PSScriptRoot '_common.ps1'
. $commonScript

# How often the parent collects worker output.
$ReceivePollMs = 400
# Silence longer than this prints a heartbeat, so a wedged run looks different from a slow one.
$HeartbeatSeconds = 10
# Shutdowns slower than this are counted separately in the summary: a clean floor close is ~2.5s.
$SlowShutdownSeconds = 10
# Settle after RegisterXLL when a suite does not say: add-in start-up is not instantaneous.
$DefaultSettleSeconds = 3

# ---------------------------------------------------------------------------
# The pid ledger. Workers are Start-Job child processes that survive the
# parent being killed, each with an Excel of its own, so every Excel is written
# to a ledger before it is used; the parent clears its ledger in a finally and
# -Cleanup sweeps every ledger for when even that did not run. One ledger per
# invocation, so two runs cannot kill each other's Excels, and it holds only
# pids we started: other people's Excels are never touched.
# ---------------------------------------------------------------------------
$ledgerDir  = Join-Path $env:TEMP 'StretchXL'
$ledgerPath = Join-Path $ledgerDir "StretchXL_pids.$PID.txt"
New-Item -ItemType Directory -Force $ledgerDir | Out-Null

function Stop-LedgerProcesses([string]$LedgerFile) {
    if (-not (Test-Path $LedgerFile)) { return 0 }
    $killed = 0
    foreach ($line in @(Get-Content $LedgerFile -ErrorAction SilentlyContinue)) {
        if ($line -notmatch '^\d+$') { continue }
        if (Stop-ExcelByPid ([int]$line)) { $killed++ }
    }
    Remove-Item $LedgerFile -Force -ErrorAction SilentlyContinue
    return $killed
}

if ($PSCmdlet.ParameterSetName -eq 'Cleanup') {
    $n = 0
    # StretchXL_pids.<pid>.txt (manager) and StretchXL_pids.standalone.<pid>.txt (TestKit)
    foreach ($stale in @(Get-ChildItem (Join-Path $ledgerDir 'StretchXL_pids.*.txt') -ErrorAction SilentlyContinue)) {
        $n += Stop-LedgerProcesses $stale.FullName
    }
    Write-Output ("StretchXL cleanup: killed {0} orphaned Excel process(es)" -f $n)
    exit 0
}

Remove-Item $ledgerPath -Force -ErrorAction SilentlyContinue

# PowerShell names are case-insensitive, so a local $runs would be the $Runs
# parameter; locals below are named so they cannot collide with one.

# the floor is the fresh-session control by definition
if ($SessionMode -ne 'Fresh' -and -not $PSBoundParameters.ContainsKey('Path')) {
    throw "-SessionMode $SessionMode needs -Path: the floor is the fresh-session control by definition"
}

New-Item -ItemType Directory -Force $OutDir | Out-Null
$outRoot = (Resolve-Path -LiteralPath $OutDir).Path
$dumpRoot = if ($PSBoundParameters.ContainsKey('DumpDir')) {
    New-Item -ItemType Directory -Force $DumpDir | Out-Null
    (Resolve-Path -LiteralPath $DumpDir).Path
} else { $outRoot }
if ($NoDump) { $dumpRoot = '' }

$workRoot = Join-Path $outRoot 'work'
New-Item -ItemType Directory -Force $workRoot | Out-Null
# Per-test stdout/stderr are captured to temp files, embedded into the JSONL,
# then deleted -- captured output travels in the results, not beside them.
$logRoot = Join-Path $ledgerDir "logs.$PID"
New-Item -ItemType Directory -Force $logRoot | Out-Null

# resolved before any Excel starts, not at the moment there is something to dump
$dumpScript = ''
if (-not $NoDump) {
    $dumpScript = Join-Path $PSScriptRoot 'evidence\hangdump.ps1'
    if (-not (Test-Path -LiteralPath $dumpScript)) {
        throw "hangdump.ps1 not found at $dumpScript -- pass -NoDump to run without evidence"
    }
}

# ---------------------------------------------------------------------------
# Discovery. Every *.test.ps1 under -Path is a test; running a subset is
# pointing at a subfolder.
# ---------------------------------------------------------------------------
$testFiles = @()
$testRoot  = ''
if ($PSBoundParameters.ContainsKey('Path')) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "-Path '$Path' does not exist" }
    $testRoot  = (Resolve-Path -LiteralPath $Path).Path
    $testFiles = @(Get-ChildItem -LiteralPath $testRoot -Recurse -Filter '*.test.ps1' |
                   Sort-Object FullName)
    if ($testFiles.Count -eq 0) {
        throw "-Path '$Path' contains no *.test.ps1 files -- nothing to run is a mistake, not an empty success"
    }
}
$floorMode = ($testFiles.Count -eq 0)

# ---------------------------------------------------------------------------
# Suite configuration: the keys and their meaning are documented in
# StretchXL.md. Resolved once per suite folder (by _common.ps1, shared with the
# kit) and validated before a single Excel starts, so a missing XLL is a
# refusal now rather than a SETUP-FAILED an hour into a soak.
# ---------------------------------------------------------------------------
$suiteConfigs = @{}   # suite dir -> resolved config
foreach ($tf in $testFiles) {
    $sd = Split-Path $tf.FullName -Parent
    if (-not $suiteConfigs.ContainsKey($sd)) {
        $sc = Resolve-SuiteConfig -SuiteDir $sd -RootDir $testRoot
        Assert-SuiteArtifacts -SuiteDir $sd -Config $sc
        $suiteConfigs[$sd] = $sc
    }
}

function Get-SessionKey($Config, [int]$Settle) {
    # Tests share a reused session only if everything that shapes the session matches.
    # RequireNotOlderThan is not part of it: it is checked before any Excel starts.
    $envPart = ''
    $sessionEnv = $Config['SessionEnvironment']
    if ($sessionEnv) {
        $envPart = (@($sessionEnv.Keys) | Sort-Object | ForEach-Object { "$_=$($sessionEnv[$_])" }) -join ';'
    }
    return ((@($Config['RegisterXll']) -join '|') + '|' + $Settle + '|' + $envPart)
}

# ---------------------------------------------------------------------------
# The work list: one item per (test, pass), or per floor iteration, built
# pass by pass so an unshuffled soak still interleaves every test each pass.
# ---------------------------------------------------------------------------
$seedUsed = $null
$workItems = New-Object System.Collections.ArrayList
if ($floorMode) {
    for ($pass = 1; $pass -le $Runs; $pass++) {
        [void]$workItems.Add(@{ Kind = 'floor'; Run = $pass; Test = ''; Rel = ''; Suite = '' })
    }
}
else {
    for ($pass = 1; $pass -le $Runs; $pass++) {
        foreach ($tf in $testFiles) {
            $relPath = $tf.FullName.Substring($testRoot.Length).TrimStart('\') -replace '\\', '/'
            $suiteRel = Split-Path $relPath -Parent
            if ([string]::IsNullOrEmpty($suiteRel)) { $suiteRel = '.' }
            $sc = $suiteConfigs[(Split-Path $tf.FullName -Parent)]
            # upward only: a suite can declare itself slow, never shorten a deadline the run set
            $effTimeout = $TestTimeoutSeconds
            if ($sc.ContainsKey('TestTimeoutSeconds')) {
                $effTimeout = [Math]::Max($effTimeout, [int]$sc['TestTimeoutSeconds'])
            }
            $settleSecs = $DefaultSettleSeconds
            if ($sc.ContainsKey('RegisterXllSettleSeconds')) { $settleSecs = [int]$sc['RegisterXllSettleSeconds'] }
            [void]$workItems.Add(@{ Kind = 'test'; Run = $pass; Test = $tf.FullName
                                    Rel = $relPath; Suite = ($suiteRel -replace '\\', '/')
                                    Xlls = @($sc['RegisterXll']); TestTimeout = $effTimeout
                                    Settle = $settleSecs; SessionEnv = $sc['SessionEnvironment']
                                    SessionKey = (Get-SessionKey $sc $settleSecs) })
        }
    }
}
if ($RandomOrder) {
    $seedUsed = if ($Seed -ge 0) { $Seed } else { Get-Random -Maximum ([int]::MaxValue) }
    $rng = New-Object System.Random($seedUsed)
    for ($i = $workItems.Count - 1; $i -gt 0; $i--) {   # Fisher-Yates
        $j = $rng.Next($i + 1)
        $tmpSwap = $workItems[$i]; $workItems[$i] = $workItems[$j]; $workItems[$j] = $tmpSwap
    }
}
$groupCount = 0
$grouped = Resolve-GroupBySession -Group:$GroupBySession -NoGroup:$NoGroupBySession -RandomOrder:$RandomOrder `
                                  -SessionMode $SessionMode -FloorMode:$floorMode
if ($grouped) {
    $workItems = Group-WorkBySession -Items $workItems -Rng $(if ($RandomOrder) { $rng } else { $null })
    $groupCount = @($workItems | ForEach-Object { $_.SessionKey } | Select-Object -Unique).Count
}

# -Parallel divides the work; the remainder goes to the first workers so the
# total is exact, and no more workers start than there are items.
$totalItems  = $workItems.Count
$workerCount = [Math]::Min($Parallel, $totalItems)
$baseShare   = [Math]::Floor($totalItems / $workerCount)
$remainder   = $totalItems % $workerCount
$shares      = @(1..$workerCount | ForEach-Object { $baseShare + $(if ($_ -le $remainder) { 1 } else { 0 }) })

# ---------------------------------------------------------------------------
# Results file: one per invocation, first line a metadata record, because a
# results file that does not say how it was produced cannot be compared.
# ---------------------------------------------------------------------------
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$resultsPath = Join-Path $outRoot "results-$stamp.jsonl"

# The Excel build, from file metadata: without it "the XLL regressed" and
# "Office updated overnight" look the same.
$excelExe = ''; $excelVer = ''
try {
    $appPathKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
    $excelExe = (Get-ItemProperty -Path $appPathKey -ErrorAction Stop).'(default)'
    if ($excelExe -and (Test-Path -LiteralPath $excelExe)) {
        $excelVer = (Get-Item -LiteralPath $excelExe).VersionInfo.FileVersion
    }
} catch {}

# The add-ins under test, by content: a RegisterXll path names a different
# file after every rebuild. File metadata only -- no COM, no load.
$productRecord = @(
    $(foreach ($sc in $suiteConfigs.Values) { @($sc['RegisterXll']) }) |
        Where-Object { $_ } | Sort-Object -Unique | ForEach-Object {
            $fi = Get-Item -LiteralPath $_
            $vi = [Diagnostics.FileVersionInfo]::GetVersionInfo($fi.FullName)
            # With no VERSIONINFO, IsDebug reads $false, which would claim
            # "release"; both fields stay empty and sha256 identifies the file.
            $hasVersionResource = -not [string]::IsNullOrWhiteSpace($vi.FileVersion)
            [ordered]@{
                path        = $fi.FullName
                name        = $fi.Name
                fileVersion = $(if ($hasVersionResource) { $vi.FileVersion } else { $null })
                isDebug     = $(if ($hasVersionResource) { [bool]$vi.IsDebug } else { $null })
                size        = $fi.Length
                modified    = $fi.LastWriteTime.ToString('o')
                sha256      = (Get-FileHash -LiteralPath $fi.FullName -Algorithm SHA256).Hash.ToLower()
            }
        }
)

$metaRecord = [ordered]@{
    type      = 'meta'
    tool      = 'StretchXL'
    started   = (Get-Date -Format 'o')
    machine   = $env:COMPUTERNAME
    excelPath = $excelExe
    excelVer  = $excelVer
    product   = @($productRecord)
    params    = [ordered]@{
        parallel = $workerCount; runs = $Runs; path = $testRoot
        sessionMode = $SessionMode; closeWithX = $CloseWithX.IsPresent
        closeTimeoutSeconds = $CloseTimeoutSeconds; testTimeoutSeconds = $TestTimeoutSeconds
        warmup = $Warmup.IsPresent; randomOrder = $RandomOrder.IsPresent; seed = $seedUsed
        groupBySession = $grouped
        dumpAfterSeconds = $DumpAfterSeconds; fullDump = $FullDump.IsPresent; noDump = $NoDump.IsPresent
        outDir = $outRoot; dumpDir = $dumpRoot
    }
    tests     = @($testFiles | ForEach-Object { $_.FullName.Substring($testRoot.Length).TrimStart('\') -replace '\\', '/' })
    suites    = @(foreach ($sd in ($suiteConfigs.Keys | Sort-Object)) {
                    $sc = $suiteConfigs[$sd]
                    if ($sc.Count -eq 0) { continue }
                    [ordered]@{
                        dir = $sd.Substring($testRoot.Length).TrimStart('\') -replace '\\', '/'
                        registerXll = @($sc['RegisterXll'])
                        testTimeoutSeconds = $sc['TestTimeoutSeconds']
                        modulesOfInterest = @($sc['ModulesOfInterest'])
                    }
                })
}
# No BOM: it would make the first record fail ordinary JSON parsers.
[System.IO.File]::WriteAllText(
    $resultsPath,
    ($metaRecord | ConvertTo-Json -Compress -Depth 8) + [Environment]::NewLine,
    (New-Object System.Text.UTF8Encoding($false)))

$workloadDesc = if ($floorMode) { 'FLOOR -- start Excel, add a new workbook, close' }
                else { "{0} test(s) under {1}" -f $testFiles.Count, $testRoot }

Write-Output 'StretchXL starting'
Write-Output ("  workload         : {0}" -f $workloadDesc)
Write-Output ("  TOTAL ITEMS      : {0}{1}" -f $totalItems, $(if (-not $floorMode) { " ($($testFiles.Count) test(s) x $Runs run(s))" } else { '' }))
Write-Output ("  parallel workers : {0}{1}" -f $workerCount,
              $(if ($workerCount -lt $Parallel) { " (asked for $Parallel; capped at the item count)" } else { '' }))
Write-Output ("  split            : {0}" -f ($shares -join ' + '))
Write-Output ("  session mode     : {0}" -f $SessionMode)
Write-Output ("  close method     : {0}" -f $(if ($CloseWithX) { 'simulated X click (visible window, WM_SYSCOMMAND SC_CLOSE)' } else { 'COM Application.Quit()' }))
Write-Output ("  close timeout    : {0}s (classifier only -- shutdown= is measured every session)" -f $CloseTimeoutSeconds)
if (-not $floorMode) {
    Write-Output ("  test timeout     : {0}s" -f $TestTimeoutSeconds)
}
if ($Warmup)      { Write-Output '  warm-up          : one unmeasured bare close per worker, reported but never counted' }
if ($RandomOrder) { Write-Output ("  order            : SHUFFLED, seed {0} (replay with -RandomOrder -Seed {0}{1})" -f $seedUsed, $(if ($NoGroupBySession -and $SessionMode -ne 'Fresh') { ' -NoGroupBySession' } else { '' })) }
if ($grouped) { Write-Output ("  grouping         : {0} session set-up(s), each kept together{1}" -f $groupCount, $(if ($RandomOrder) { '; group order shuffled by the seed' } else { '' })) }
Write-Output ("  dump on hang     : {0}" -f $(if ($NoDump) { 'NO -- -NoDump; hangs are killed with no evidence' } elseif ($FullDump) { 'yes, full memory' } else { 'yes, stacks and handles' }))
if ($DumpAfterSeconds -gt 0) {
    Write-Output ("  dump if slow     : after {0}s still alive, while still stuck" -f $DumpAfterSeconds)
}
Write-Output ("  results          : {0}" -f $resultsPath)
# the binaries under test on stdout too, so the scrollback says what was tested
foreach ($p in $productRecord) {
    $what = if ($null -eq $p.isDebug) { '(no version resource)' }
            elseif ($p.isDebug)       { "v$($p.fileVersion) DEBUG BUILD" }
            else                      { "v$($p.fileVersion) release" }
    Write-Output ("  under test       : {0,-18} {1,-24} sha {2}" -f `
                  $p.name, $what, $p.sha256.Substring(0, 12))
}
if ($productRecord | Where-Object { $_.isDebug }) {
    Write-Output '  NOTE             : a DEBUG build is under test -- valid for debugging, never for a release'
}
Write-Output ''

# ---------------------------------------------------------------------------
# One worker: runs its share of items and emits one 'RESULT {json}' line per
# item, returned rather than printed so the parent owns ordering and counting.
# Every path emits a line: "no line" and "nothing happened" must never look
# the same.
#
# Each item is one session bracket: start Excel, identify it by its own
# window, ledger the pid, open our own process handle, add a workbook, run the
# test, close the configured way, release every reference, wait on the
# handle, classify.
# ---------------------------------------------------------------------------
$workerBody = {
    param($Cfg)

    $ErrorActionPreference = 'Stop'
    $WorkerId   = $Cfg.WorkerId
    $LedgerFile = $Cfg.Ledger

    # A job is its own process, so the shared helpers are loaded by path.
    if (-not (Test-Path -LiteralPath $Cfg.CommonScript)) {
        throw "StretchXL worker ${WorkerId}: shared helpers not found at '$($Cfg.CommonScript)'"
    }
    . $Cfg.CommonScript

    # Retries: 20 tries at the delays below outlast contention without stalling an item.
    $RetryTries = 20
    # Several workers append to the one ledger, so a write can meet another's lock.
    $LedgerRetryDelayMs = 50
    # The watchdog reads and appends its files every poll, so an open can collide briefly.
    $WatchFileRetryDelayMs = 25
    # Excel's Hwnd is not always readable the instant its COM object exists.
    $WindowLookupDelayMs = 100
    # A killed test process exits almost at once; this only bounds that wait.
    $KilledTestExitWaitMs = 5000
    # Polls the pid while waiting on a close with no process handle.
    $ClosePollMs = 250
    # Enough of a test's stdout for its protocol lines and a useful log tail; stderr needs less.
    $TestLogTailBytes = 8192
    $TestErrTailBytes = 4096
    # A dialog's text is cut to this on a result line; the full text is in the dialog log.
    $DialogTextMaxChars = 160
    # Screenshots larger than this are not real windows.
    $ShotMaxPixels = 8192

    # Record the pid before using the process, so a worker killed a moment
    # later still leaves its Excel findable.
    function Add-ToLedger([int]$ExcelPid, [string]$File) {
        Invoke-WithRetry -Tries $RetryTries -DelayMs $LedgerRetryDelayMs -Action {
            Add-Content -Path $File -Value $ExcelPid -Encoding ascii -ErrorAction Stop
        }
    }

    # Dump a process while it is still stuck. Failure returns '': no dump is a
    # fact the result states, never an exception that loses the hang.
    function Get-HangDump([int]$TargetPid, [string]$Label) {
        if ([string]::IsNullOrEmpty($Cfg.DumpScript) -or [string]::IsNullOrEmpty($Cfg.DumpDir)) { return '' }
        $dumpFile = Join-Path $Cfg.DumpDir ("{0}.{1}.{2}.hang.dmp" -f $Label, $TargetPid, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        try {
            $dumpResult = if ($Cfg.FullDump) { & $Cfg.DumpScript -ProcessId $TargetPid -Path $dumpFile }
                          else               { & $Cfg.DumpScript -ProcessId $TargetPid -Path $dumpFile -Small }
            $last = @($dumpResult) | Select-Object -Last 1
            if ($last -and $last.Path) { return [string]$last.Path }
        } catch {}
        return ''
    }

    # P/Invoke rather than System.Diagnostics.Process: Process.ExitCode reads
    # $null for a process the object did not start, which would count every
    # crash as clean. Our own handle gives the exit code, a real wait, and
    # pins the pid against recycling.
    Add-Type -Name Win -Namespace SX -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern bool PostMessage(System.IntPtr h, uint msg, System.IntPtr w, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr h, int cmd);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern System.IntPtr OpenProcess(uint access, bool inherit, int pid);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool GetExitCodeProcess(System.IntPtr h, out uint code);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern uint WaitForSingleObject(System.IntPtr h, uint ms);
[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError=true)]
public static extern bool CloseHandle(System.IntPtr h);
public delegate bool WndEnumProc(System.IntPtr h, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool EnumWindows(WndEnumProc cb, System.IntPtr l);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool IsWindowVisible(System.IntPtr h);
[System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool GetWindowRect(System.IntPtr h, out RECT r);
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool PrintWindow(System.IntPtr h, System.IntPtr hdc, uint flags);
'@

    # Screenshot before killing: a hidden Excel's modal dialog is still its own
    # visible window, and it shows the question a kill would destroy.
    function Save-WindowShots([int]$TargetPid, [string]$Label) {
        if ([string]::IsNullOrEmpty($Cfg.DumpDir)) { return @() }
        $saved = @()
        try {
            Add-Type -AssemblyName System.Drawing -ErrorAction Stop
            $wins = New-Object System.Collections.ArrayList
            $enumCb = [SX.Win+WndEnumProc]{ param($h, $l)
                $p = 0; [void][StretchXL.WindowProcess]::GetWindowThreadProcessId($h, [ref]$p)
                if ($p -eq $TargetPid -and [SX.Win]::IsWindowVisible($h)) { [void]$wins.Add($h) }
                return $true }
            [void][SX.Win]::EnumWindows($enumCb, [IntPtr]::Zero)
            $shotIdx = 0
            foreach ($winH in $wins) {
                $rc = New-Object 'SX.Win+RECT'
                if (-not [SX.Win]::GetWindowRect($winH, [ref]$rc)) { continue }
                $wpx = $rc.Right - $rc.Left; $hpx = $rc.Bottom - $rc.Top
                if ($wpx -le 0 -or $hpx -le 0 -or $wpx -gt $ShotMaxPixels -or $hpx -gt $ShotMaxPixels) { continue }
                $bmp = New-Object System.Drawing.Bitmap($wpx, $hpx)
                $gfx = [System.Drawing.Graphics]::FromImage($bmp)
                $hdc = $gfx.GetHdc()
                $okShot = [SX.Win]::PrintWindow($winH, $hdc, 2)   # PW_RENDERFULLCONTENT
                $gfx.ReleaseHdc($hdc); $gfx.Dispose()
                if ($okShot) {
                    $shotIdx++
                    $shotFile = Join-Path $Cfg.DumpDir ("{0}.{1}.w{2}.{3}.png" -f `
                                $Label, $TargetPid, $shotIdx, (Get-Date -Format 'yyyyMMdd-HHmmss'))
                    $bmp.Save($shotFile, [System.Drawing.Imaging.ImageFormat]::Png)
                    $saved += $shotFile
                }
                $bmp.Dispose()
            }
        } catch {}
        return $saved
    }
    # SYNCHRONIZE | PROCESS_QUERY_LIMITED_INFORMATION: wait and read the exit
    # code only. Killing goes through the named, pid-checked path.
    $PROC_ACCESS  = 0x00100000 -bor 0x00001000
    $WAIT_OBJECT0 = 0
    $STILL_ACTIVE = 259
    # What the X button sends. Posted, never sent: SendMessage would block
    # this worker inside the hang it is measuring.
    $WM_SYSCOMMAND = [uint32]0x0112
    $SC_CLOSE      = [IntPtr]0xF060

    # Our Excel's window and pid from its own Hwnd, never by image name: other
    # people's Excels are on this machine and parallel workers would race.
    function Get-OwnExcelWindow($app) {
        Invoke-WithRetry -Tries $RetryTries -DelayMs $WindowLookupDelayMs -Action {
            $rawHwnd = [int64]$app.Hwnd
            $found = if ($rawHwnd -ne 0) { Get-WindowProcessId ([IntPtr]$rawHwnd) } else { 0 }
            if ($found -eq 0) { throw 'Excel window not ready' }
            @{ Hwnd = [IntPtr]$rawHwnd; OwnPid = $found }
        }
    }

    function Read-Tail([string]$File, [int]$MaxBytes) {
        if (-not (Test-Path -LiteralPath $File)) { return '' }
        try {
            $raw = [System.IO.File]::ReadAllText($File)
            if ($raw.Length -gt $MaxBytes) { $raw = $raw.Substring($raw.Length - $MaxBytes) }
            return $raw
        } catch { return '' }
    }

    # ------------------------------------------------------------------------
    # The modal-dialog watchdog: dismiss and fail. Excel blocks the calling
    # thread on a modal dialog, so a MsgBox or VBA error would hold a session
    # until its deadline. One watchdog job per worker, told which pid to watch
    # through a file; it touches that pid's dialogs only, logs each dismissal,
    # and the manager fails a PASS that carries an unacknowledged dialog.
    # ------------------------------------------------------------------------
    $dlgPidFile = Join-Path $Cfg.LogDir ("w{0}.watchpid" -f $WorkerId)
    $dlgLogFile = Join-Path $Cfg.LogDir ("w{0}.dialogs.log" -f $WorkerId)
    # File.WriteAllText shares with the watchdog's reads, where Set-Content would not.
    function Set-WatchPid([string]$Value) {
        Invoke-WithRetry -Tries $RetryTries -DelayMs $WatchFileRetryDelayMs -Action {
            [System.IO.File]::WriteAllText($dlgPidFile, $Value)
        }
    }
    Set-WatchPid '0'
    $watchdogJob = Start-Job -ArgumentList $dlgPidFile, $dlgLogFile, $PID, $Cfg.CommonScript {
        param($pidFile, $logFile, $ownerPid, $commonPath)
        if (-not (Test-Path -LiteralPath $commonPath)) { throw "StretchXL watchdog: shared helpers not found at '$commonPath'" }
        . $commonPath
        # Often enough that a dialog is gone well inside any deadline, rarely enough to cost nothing.
        $DialogPollMs = 400
        Add-Type @"
using System;using System.Text;using System.Runtime.InteropServices;
public class SXDW {
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr h, EnumWindowsProc cb, IntPtr l);
  public delegate bool EnumWindowsProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
}
"@
        while ($true) {
            # exit with the worker, or this becomes an orphan the parent cannot see
            if (-not (Get-Process -Id $ownerPid -ErrorAction SilentlyContinue)) { return }
            $watchPid = 0
            try { $watchPid = [int](Get-Content $pidFile -ErrorAction SilentlyContinue | Select-Object -First 1) } catch {}
            if ($watchPid -gt 0) {
                $dlgs = New-Object System.Collections.ArrayList
                $cb = [SXDW+EnumWindowsProc]{ param($h, $l)
                    $p = 0; [void][StretchXL.WindowProcess]::GetWindowThreadProcessId($h, [ref]$p)
                    if ($p -eq $watchPid -and [SXDW]::IsWindowVisible($h)) {
                        $c = New-Object System.Text.StringBuilder 256
                        [void][SXDW]::GetClassName($h, $c, 256)
                        if ($c.ToString() -eq '#32770') { [void]$dlgs.Add($h) }
                    }
                    return $true }
                [void][SXDW]::EnumWindows($cb, [IntPtr]::Zero)
                foreach ($d in $dlgs) {
                    $texts = New-Object System.Collections.ArrayList
                    $btns  = @{}
                    $cb2 = [SXDW+EnumWindowsProc]{ param($h, $l)
                        $t = New-Object System.Text.StringBuilder 512
                        [void][SXDW]::GetWindowText($h, $t, 512)
                        $c = New-Object System.Text.StringBuilder 128
                        [void][SXDW]::GetClassName($h, $c, 128)
                        $s = $t.ToString()
                        if ($s) {
                            if ($c.ToString() -eq 'Button') { $btns[($s -replace '&', '')] = $h }
                            else { [void]$texts.Add($s) }
                        }
                        return $true }
                    [void][SXDW]::EnumChildWindows($d, $cb2, [IntPtr]::Zero)
                    $msg = ($texts -join ' / ') -replace "`r`n", ' '
                    $hit = $null
                    # "End" first: on a VBA runtime-error dialog, Continue or
                    # Debug leave state the next case inherits.
                    foreach ($pref in @('End', 'OK', 'Yes', 'Close', 'Cancel', 'No')) {
                        if ($btns.ContainsKey($pref)) { $hit = $btns[$pref]; break }
                    }
                    if ($hit) {
                        [void][SXDW]::SendMessage($hit, 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)  # BM_CLICK
                        Add-Content -Path $logFile -Value ("DIALOG: " + $msg)
                    }
                }
            }
            Start-Sleep -Milliseconds $DialogPollMs
        }
    }
    # The watchdog appends while we read: share ReadWrite and retry, since a
    # failed read would look like "no dialogs".
    function Get-DialogLines {
        if (-not (Test-Path -LiteralPath $dlgLogFile)) { return @() }
        @(Invoke-WithRetry -Tries $RetryTries -DelayMs $WatchFileRetryDelayMs -Action {
            $fs = [System.IO.File]::Open($dlgLogFile, 'Open', 'Read', 'ReadWrite')
            try {
                $sr = New-Object System.IO.StreamReader($fs)
                $text = $sr.ReadToEnd()
                $sr.Dispose()
            } finally { $fs.Dispose() }
            if (-not [string]::IsNullOrEmpty($text)) { $text -split "`r?`n" | Where-Object { $_ -ne '' } }
        })
    }
    function Get-DialogLineCount { @(Get-DialogLines).Count }

    # " dialogs=N last='...'" for a result's detail.
    function Format-DialogDetail($Lines) {
        $lastDlg = ($Lines[-1] -replace '^DIALOG:\s*', '')
        if ($lastDlg.Length -gt $DialogTextMaxChars) { $lastDlg = $lastDlg.Substring(0, $DialogTextMaxChars) }
        " dialogs=$($Lines.Count) last='$lastDlg'"
    }

    # Wall clock since this worker started, on every result as t=: a failure at
    # a fixed iteration and one at a fixed elapsed time are different mechanisms.
    $workerClock = [Diagnostics.Stopwatch]::StartNew()
    $iterIndex   = 0

    # The warm-up is item 0: a bare floor close through the same machinery,
    # paying this worker's cold interactive-close cost outside the data.
    $itemQueue = New-Object System.Collections.ArrayList
    if ($Cfg.Warmup) { [void]$itemQueue.Add(@{ Kind = 'warmup'; Run = 0; Test = ''; Rel = ''; Suite = '' }) }
    foreach ($it in $Cfg.Items) { [void]$itemQueue.Add($it) }

    # ------------------------------------------------------------------------
    # Session state, held across items. Fresh never keeps a session past its
    # item; Reuse and ReuseClean keep one for consecutive items with the same
    # SessionKey, so one session never mixes set-ups, and folder boundaries with
    # the same set-up do not cycle it.
    # ------------------------------------------------------------------------
    $script:sApp = $null; $script:sBooks = $null; $script:sBook = $null
    $script:sHandle = $null; $script:sPid = 0; $script:sWin = $null
    $script:sSeq = 0; $script:sEnvNames = @(); $script:sDir = ''; $script:sBaseBook = ''
    $script:sKey = ''; $script:sOrdinal = 0; $script:sServed = 0
    $script:sWasTest = $false

    # The exit code from our own handle, or $null when it cannot be read.
    function Read-SessionExitCode {
        if (-not $script:sHandle) { return $null }
        $raw = [uint32]0
        # STILL_ACTIVE means the process has not exited, not that it returned 259
        if ([SX.Win]::GetExitCodeProcess($script:sHandle, [ref]$raw) -and $raw -ne $STILL_ACTIVE) { return $raw }
        return $null
    }

    function Close-LiveSession {
        # The measured close, one implementation for every mode. Returns the
        # facts and clears the session state.
        $r = @{ Outcome = ''; ShutSecs = $null; Detail = ''; Dumps = @(); Shots = @(); DlgLines = @() }
        if ($null -eq $script:sApp -and $script:sPid -eq 0) { return $r }
        $closeDlgMark = Get-DialogLineCount
        $quitClock = $null

        try {
            # Test sessions only: a dirty workbook turns the close into a hidden
            # "Save changes?" dialog, and the manager's close means discard. The
            # floor never dirties a book and keeps its exact COM sequence.
            if ($script:sWasTest) {
                try {
                    $sweepBooks = $script:sApp.Workbooks
                    for ($bi = 1; $bi -le $sweepBooks.Count; $bi++) {
                        $wbRef = $sweepBooks.Item($bi)
                        try { $wbRef.Saved = $true } catch {}
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbRef)
                    }
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sweepBooks)
                } catch {}
            }

            # X-arm setup before the clock: arranging the scenario is not the
            # shutdown being measured. Minimised without activation, then visible.
            if ($Cfg.CloseWithX) {
                [void][SX.Win]::ShowWindow($script:sWin.Hwnd, 7)   # SW_SHOWMINNOACTIVE
                $script:sApp.Visible = $true
            }

            $quitClock = [Diagnostics.Stopwatch]::StartNew()
            if ($Cfg.CloseWithX) {
                if (-not [SX.Win]::PostMessage($script:sWin.Hwnd, $WM_SYSCOMMAND, $SC_CLOSE, [IntPtr]::Zero)) {
                    $r.Detail = ("close=POST-FAILED " + $r.Detail).Trim()
                }
            }
            else {
                $script:sApp.Quit()
            }

            # Release every reference while Excel is still there to receive it:
            # an unreleased reference manufactures the hang this harness detects.
            foreach ($comRef in @($script:sBook, $script:sBooks, $script:sApp)) {
                if ($comRef) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comRef) } catch {} }
            }
        }
        catch {
            $r.Detail = ('automation=' + ($_.Exception.Message -replace '\s+', ' ') + ' ' + $r.Detail).Trim()
        }
        finally {
            $script:sBook = $null; $script:sBooks = $null; $script:sApp = $null
        }

        # A bounded wait on our own handle, two-staged when -DumpAfterSeconds
        # is set so the dump is taken while still stuck.
        $exited = $false
        $closePid = $script:sPid
        if ($null -eq $script:sHandle) {
            $deadline = (Get-Date).AddSeconds($Cfg.CloseTimeout)
            while ((Get-Date) -lt $deadline -and (Get-Process -Id $closePid -ErrorAction SilentlyContinue)) {
                Start-Sleep -Milliseconds $ClosePollMs
            }
            $exited = -not (Get-Process -Id $closePid -ErrorAction SilentlyContinue)
        }
        else {
            if ($Cfg.DumpAfter -gt 0 -and $Cfg.DumpAfter -lt $Cfg.CloseTimeout) {
                $exited = ([SX.Win]::WaitForSingleObject($script:sHandle, [uint32]($Cfg.DumpAfter * 1000)) -eq $WAIT_OBJECT0)
                if (-not $exited) {
                    $slowDump = Get-HangDump $closePid 'EXCEL'
                    if ($slowDump) { $r.Dumps += $slowDump }
                }
            }
            if (-not $exited) {
                # measured from the close, not from the dump
                $left = [int](($Cfg.CloseTimeout * 1000) - $quitClock.Elapsed.TotalMilliseconds)
                if ($left -lt 0) { $left = 0 }
                $exited = ([SX.Win]::WaitForSingleObject($script:sHandle, [uint32]$left) -eq $WAIT_OBJECT0)
            }
        }
        if ($quitClock) { $quitClock.Stop(); $r.ShutSecs = [math]::Round($quitClock.Elapsed.TotalSeconds, 1) }

        if (-not $exited) {
            # A hang is a result: dumped and screenshotted before anything is killed.
            $r.Outcome = 'HANG'
            $title = ''
            try { $title = (Get-Process -Id $closePid -ErrorAction SilentlyContinue).MainWindowTitle } catch {}
            $hangDump = Get-HangDump $closePid 'EXCEL'
            if ($hangDump) { $r.Dumps += $hangDump }
            $r.Shots += @(Save-WindowShots $closePid 'EXCEL')
            $r.Detail = ("waited=$($Cfg.CloseTimeout)s window='$title' " + $r.Detail).Trim()
            if ($r.Dumps.Count -eq 0 -and -not [string]::IsNullOrEmpty($Cfg.DumpScript)) {
                $r.Detail += ' dump=FAILED'
            }
            [void](Stop-ExcelByPid $closePid)
        }
        else {
            $code = Read-SessionExitCode
            if ($null -eq $code) {
                # never silently clean
                $r.Outcome = 'EXITED'
                $r.Detail  = ('exit=unknown ' + $r.Detail).Trim()
            }
            elseif ($code -ne 0) {
                $r.Outcome = 'CRASH'
                $r.Detail  = (('exit=0x{0:X8} ' -f [uint32]$code) + $r.Detail).Trim()
            }
            else { $r.Outcome = 'CLEAN' }
        }

        if ($script:sHandle) { [void][SX.Win]::CloseHandle($script:sHandle); $script:sHandle = $null }
        Set-WatchPid '0'
        $dlgNow = Get-DialogLineCount
        if ($dlgNow -gt $closeDlgMark) {
            $allDlg = @(Get-DialogLines)
            $r.DlgLines = @($allDlg[$closeDlgMark..($dlgNow - 1)])
        }
        $script:sPid = 0; $script:sWin = $null; $script:sKey = ''; $script:sWasTest = $false
        return $r
    }

    # A reused session's close is its own result: shutdown= measured once per
    # session, with the count of tests it served.
    function Close-SessionWithRecord {
        $servedCount = $script:sServed
        $closedPid = $script:sPid
        $closeR = Close-LiveSession
        $detail = $closeR.Detail
        if ($closeR.DlgLines.Count -gt 0) { $detail = ($detail + (Format-DialogDetail $closeR.DlgLines)).Trim() }
        $rec = [ordered]@{
            type = 'result'; kind = 'session-close'; worker = $WorkerId; run = $null
            test = $null; suite = $null; pid = $closedPid
            session = $script:sOrdinal; spos = $null; served = $servedCount
            verdict = $null; outcome = $closeR.Outcome
            shutdown_s = $closeR.ShutSecs; test_s = $null
            elapsed_s = $closeR.ShutSecs
            t_s = [math]::Round($workerClock.Elapsed.TotalSeconds, 1)
            cases_pass = 0; cases_fail = 0
            detail = $detail; dumps = @($closeR.Dumps); shots = @($closeR.Shots)
            dialogs = $closeR.DlgLines.Count; test_exit = $null; log = ''; stderr = ''
        }
        'RESULT ' + ($rec | ConvertTo-Json -Compress -Depth 4)
    }

    # ---- item phases. Each takes the item and its state table and fills it in.

    function New-ItemState($item) {
        $isTest = ($item.Kind -eq 'test')
        @{
            Clock = [Diagnostics.Stopwatch]::StartNew(); DlgMark = (Get-DialogLineCount)
            IsTest = $isTest
            # floor and warm-up items always get a fresh, bare session: the floor is the control
            Reuse = (($Cfg.SessionMode -ne 'Fresh') -and $isTest)
            Key = $(if ($isTest) { $item.SessionKey } else { '<floor>' })
            Detail = ''; Dumps = @(); Shots = @(); ShutSecs = $null
            Verdict = $null; TestSecs = $null; CasesPass = 0; CasesFail = 0
            TestLog = ''; TestErr = ''; TestExit = $null
            Spos = $null; Outcome = ''; DlgEnd = 0; ItemPid = 0; DlgLines = @()
        }
    }

    # 1. Close a lingering session this item cannot use.
    function Close-SessionIfUnusable($st) {
        if (-not ($script:sApp -or $script:sPid -ne 0)) { return }
        $stillAlive = ($script:sPid -ne 0) -and (Get-Process -Id $script:sPid -ErrorAction SilentlyContinue)
        if ($st.Reuse -and $script:sKey -eq $st.Key -and $stillAlive) { return }
        Close-SessionWithRecord
    }

    # 2. Open a session. Deliberately absent: Visible=false (COM-created Excel
    # starts hidden), DisplayAlerts=false (it suppresses the dialogs that are
    # evidence). The empty workbook stays even for tests: the kit binds through
    # its EXCEL7 window.
    function Open-Session($item, $st) {
        # A COM-created Excel inherits this process's environment, and so does
        # the test; names one session set are cleared before the next.
        foreach ($name in $script:sEnvNames) { [Environment]::SetEnvironmentVariable($name, $null, 'Process') }
        $script:sEnvNames = @()
        $script:sSeq++
        $script:sDir = Join-Path $Cfg.WorkDir ("session_w{0}_{1}" -f $Cfg.WorkerId, $script:sSeq)
        New-Item -ItemType Directory -Force $script:sDir | Out-Null
        $env:STRETCH_SESSION_DIR = $script:sDir
        if ($item.SessionEnv) {
            foreach ($name in @($item.SessionEnv.Keys)) {
                $val = ([string]$item.SessionEnv[$name]).Replace('{SessionDir}', $script:sDir)
                [Environment]::SetEnvironmentVariable($name, $val, 'Process')
                $script:sEnvNames += $name
            }
        }
        try {
            $script:sApp = New-Object -ComObject Excel.Application
            $script:sWin = Get-OwnExcelWindow $script:sApp
            if ($script:sWin) {
                $script:sPid = $script:sWin.OwnPid
                Add-ToLedger $script:sPid $LedgerFile
                # Opened now, not at the close: later it would race a fast crash
                # and there would be no exit code.
                $script:sHandle = [SX.Win]::OpenProcess($PROC_ACCESS, $false, $script:sPid)
                if ($script:sHandle -eq [IntPtr]::Zero) { $script:sHandle = $null }
                Set-WatchPid ([string]$script:sPid)
            }
            $script:sBooks = $script:sApp.Workbooks
            $script:sBook  = $script:sBooks.Add()
            # the baseline's name, so ReuseClean can close every other book
            try { $script:sBaseBook = [string]$script:sBook.Name } catch { $script:sBaseBook = '' }
            $script:sKey = $st.Key
            $script:sOrdinal++
            $script:sServed = 0
            $script:sWasTest = $st.IsTest

            # Suite set-up once per session, before any test sees it. A failure
            # is SETUP-FAILED, so a config mistake never looks like a test failure.
            if ($st.IsTest -and $script:sPid -ne 0 -and $item.Xlls -and @($item.Xlls).Count -gt 0) {
                $setupErr = ''
                foreach ($xllPath in @($item.Xlls)) {
                    if (-not $xllPath) { continue }
                    $regOk = $false
                    try { $regOk = [bool]$script:sApp.RegisterXLL([string]$xllPath) }
                    catch { $setupErr = ($_.Exception.Message -replace '\s+', ' ') }
                    if (-not $regOk) {
                        $setupErr = ("RegisterXLL failed: {0} {1}" -f $xllPath, $setupErr).Trim()
                        break
                    }
                }
                if ($setupErr) { $st.Verdict = 'SETUP-FAILED'; $st.Detail = $setupErr }
                else { Start-Sleep -Seconds $item.Settle }
            }
        }
        catch {
            $st.Detail = ('automation=' + ($_.Exception.Message -replace '\s+', ' ') + ' ' + $st.Detail).Trim()
        }
    }

    # ReuseClean between tests: back to set-up-complete, not virgin Excel (the
    # XLLs stay registered). Best-effort.
    function Reset-SessionToBaseline {
        # The baseline book stays open: the kit binds through its window, and a
        # replacement book would have a new hwnd. Every RCW touched is released,
        # or it keeps Excel alive and turns the session's Quit into a hang.
        try {
            try { $script:sApp.DisplayAlerts = $false } catch {}
            $keep = $script:sBaseBook
            $wbs = $script:sApp.Workbooks
            $wi = 1
            while ($wi -le $wbs.Count) {
                $wbRef = $wbs.Item($wi)
                $nm = ''
                try { $nm = [string]$wbRef.Name } catch {}
                if ($nm -and $nm -ne $keep) {
                    try { $wbRef.Close($false) } catch {}
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbRef)
                    # the collection shrank -- do not advance
                }
                else {
                    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbRef)
                    $wi++
                }
            }
            # Count before releasing: every mention of $app.Workbooks makes a new reference.
            $remaining = $wbs.Count
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wbs)
            # only if a test closed the baseline itself: make a new one and re-derive the hwnd
            if ($remaining -eq 0) {
                $script:sBook = $script:sBooks.Add()
                try { $script:sBaseBook = [string]$script:sBook.Name } catch {}
                $reWin = Get-OwnExcelWindow $script:sApp
                if ($reWin) { $script:sWin = $reWin }
            }
        } catch {}
        try { $script:sApp.DisplayAlerts = $true } catch {}
        try { $script:sApp.Calculation = -4105 } catch {}   # xlCalculationAutomatic
        try { $script:sApp.ScreenUpdating = $true } catch {}
        try { $script:sApp.EnableEvents = $true } catch {}
        try { $script:sApp.StatusBar = $false } catch {}
    }

    # The pid is unknown but the COM object may be live: quit and release, or
    # the held references keep an Excel the ledger and -Cleanup cannot see.
    function Write-UntrackedSessionRecord($item, $st) {
        $st.Clock.Stop()
        try { if ($script:sApp) { $script:sApp.Quit() } } catch {}
        foreach ($comRef in @($script:sBook, $script:sBooks, $script:sApp)) {
            if ($comRef) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comRef) } catch {} }
        }
        $failRec = [ordered]@{
            type = 'result'; kind = $item.Kind; worker = $WorkerId; run = $item.Run
            test = $item.Rel; suite = $item.Suite; pid = 0
            session = $script:sOrdinal; spos = $null; served = $null
            verdict = $(if ($st.IsTest) { 'ERROR' } else { $null })
            outcome = 'FAILED'; shutdown_s = $null; test_s = $st.TestSecs
            elapsed_s = [math]::Round($st.Clock.Elapsed.TotalSeconds, 1)
            t_s = [math]::Round($workerClock.Elapsed.TotalSeconds, 1)
            cases_pass = $st.CasesPass; cases_fail = $st.CasesFail
            detail = ('UNTRACKED-EXCEL-MAY-REMAIN ' + $st.Detail).Trim()
            dumps = @(); shots = @(); dialogs = 0
            test_exit = $st.TestExit; log = $st.TestLog; stderr = $st.TestErr
        }
        'RESULT ' + ($failRec | ConvertTo-Json -Compress -Depth 4)
        $script:sApp = $null; $script:sBooks = $null; $script:sBook = $null; $script:sKey = ''
    }

    # 3. The test process, against the live session.
    function Invoke-TestProcess($item, $st, [int]$Index) {
        $outLog = Join-Path $Cfg.LogDir ("w{0}-i{1}.out.txt" -f $WorkerId, $Index)
        $errLog = Join-Path $Cfg.LogDir ("w{0}-i{1}.err.txt" -f $WorkerId, $Index)

        # The contract: environment in, stdout out. Items run one at a time in
        # a worker, so setting and clearing around the spawn is race-free.
        $env:STRETCH_SESSION_HWND = [string][int64]$script:sWin.Hwnd
        $env:STRETCH_SESSION_PID  = [string]$script:sPid
        $env:STRETCH_WORKDIR      = $Cfg.WorkDir
        $env:STRETCH_LEDGER       = $LedgerFile
        $env:STRETCH_DIALOG_LOG   = $dlgLogFile
        # the test's scratch is named after the test, since pids are recycled
        $env:STRETCH_TEST_ID      = [IO.Path]::GetFileNameWithoutExtension($item.Test) -replace '\.test$',''
        $testClock = [Diagnostics.Stopwatch]::StartNew()
        $tp = $null
        # PS 5.1 Start-Process joins -ArgumentList with spaces and quotes nothing
        $fileArg = '"' + $item.Test + '"'
        try {
            $tp = Start-Process -FilePath 'powershell.exe' `
                    -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $fileArg) `
                    -RedirectStandardOutput $outLog -RedirectStandardError $errLog `
                    -PassThru -WindowStyle Hidden
        } catch {
            $st.Verdict = 'ERROR'
            $st.Detail  = ('spawn=' + ($_.Exception.Message -replace '\s+', ' '))
        }
        finally {
            Remove-Item Env:STRETCH_SESSION_HWND, Env:STRETCH_SESSION_PID,
                        Env:STRETCH_WORKDIR, Env:STRETCH_LEDGER,
                        Env:STRETCH_DIALOG_LOG, Env:STRETCH_TEST_ID `
                        -ErrorAction SilentlyContinue
        }
        if (-not $tp) { return }

        # On a TIMEOUT both processes are dumped and the session's windows
        # screenshotted before the kill: a dialog the watchdog missed would show.
        $itemDeadline = [int]$item.TestTimeout
        $finished = $tp.WaitForExit($itemDeadline * 1000)
        $testClock.Stop()
        $st.TestSecs = [math]::Round($testClock.Elapsed.TotalSeconds, 1)
        if (-not $finished) {
            $st.Verdict = 'TIMEOUT'
            $dExcel = Get-HangDump $script:sPid 'EXCEL'
            $dTest  = Get-HangDump $tp.Id 'TESTPS'
            $st.Dumps += @(@($dExcel, $dTest) | Where-Object { $_ })
            $st.Shots += @(Save-WindowShots $script:sPid 'EXCEL')
            try { $tp.Kill() } catch {}
            [void]$tp.WaitForExit($KilledTestExitWaitMs)
            $st.Detail = "test-deadline=${itemDeadline}s"
        }
        else {
            # valid here: this Process object started the process
            $st.TestExit = $tp.ExitCode
        }

        $st.TestLog = Read-Tail $outLog $TestLogTailBytes
        $st.TestErr = Read-Tail $errLog $TestErrTailBytes
        Remove-Item $outLog, $errLog -Force -ErrorAction SilentlyContinue

        # A nonzero exit, or no verdict line, is an ERROR: a test that could
        # not say what it concluded concluded nothing.
        foreach ($ln in ($st.TestLog -split "`r?`n")) {
            if ($ln -match '^STRETCH case=(.+?) verdict=(PASS|FAIL)\b') {
                if ($Matches[2] -eq 'PASS') { $st.CasesPass++ } else { $st.CasesFail++ }
            }
        }
        if ($null -ne $st.Verdict) { return }
        $verdictLine = $null
        foreach ($ln in ($st.TestLog -split "`r?`n")) {
            if ($ln -match '^STRETCH verdict=(\w+)(?: detail=(.*))?$') { $verdictLine = $Matches }
        }
        if ($null -ne $st.TestExit -and $st.TestExit -ne 0) {
            $st.Verdict = 'ERROR'
            $st.Detail  = ("exit={0} " -f $st.TestExit) + $(if ($st.TestErr) { 'stderr-captured' } else { '' })
        }
        elseif ($null -eq $verdictLine) {
            $st.Verdict = 'ERROR'
            $st.Detail  = 'no-verdict-line'
        }
        else {
            $st.Verdict = $verdictLine[1]
            if ($verdictLine.Count -gt 2 -and $verdictLine[2]) {
                $st.Detail = ($verdictLine[2] -replace '\s+', ' ').Trim()
            }
        }
    }

    # 4. End of item: close the session (fresh), or check it survived (reuse).
    function Complete-ItemSession($st) {
        if (-not $st.Reuse) {
            $st.ItemPid = $script:sPid
            $closeR = Close-LiveSession
            $st.ShutSecs = $closeR.ShutSecs
            $st.Outcome  = $closeR.Outcome
            $st.Dumps   += @($closeR.Dumps)
            $st.Shots   += @($closeR.Shots)
            if ($closeR.Detail) { $st.Detail = ($closeR.Detail + ' ' + $st.Detail).Trim() }
            # fresh: dialogs through the close belong to the item
            $st.DlgEnd = Get-DialogLineCount
            return
        }
        # Death mid-test is this item's outcome; the next item opens a replacement session.
        $st.DlgEnd = Get-DialogLineCount
        $st.ItemPid = $script:sPid
        if (Get-Process -Id $script:sPid -ErrorAction SilentlyContinue) { $st.Outcome = 'ALIVE'; return }
        $code = Read-SessionExitCode
        if ($script:sHandle) { [void][SX.Win]::CloseHandle($script:sHandle); $script:sHandle = $null }
        if ($null -ne $code -and $code -ne 0) {
            $st.Outcome = 'CRASH'
            $st.Detail  = (('session died during test: exit=0x{0:X8} ' -f [uint32]$code) + $st.Detail).Trim()
        }
        else {
            $st.Outcome = 'EXITED'
            $st.Detail  = ('session died during test (exit=' + $(if ($null -eq $code) { 'unknown' } else { '0x0' }) + ') ' + $st.Detail).Trim()
        }
        Set-WatchPid '0'
        $script:sApp = $null; $script:sBooks = $null; $script:sBook = $null
        $script:sPid = 0; $script:sWin = $null; $script:sKey = ''; $script:sWasTest = $false
    }

    # Dialogs: dismissed, recorded and failed, unless the test declared
    # 'STRETCH dialogs=handled'; the count lands on the record either way.
    function Resolve-ItemDialogs($st) {
        if ($st.DlgEnd -gt $st.DlgMark) {
            $allDlg = @(Get-DialogLines)
            $st.DlgLines = @($allDlg[$st.DlgMark..($st.DlgEnd - 1)])
        }
        if ($st.DlgLines.Count -eq 0) { return }
        $st.Detail = ($st.Detail + (Format-DialogDetail $st.DlgLines)).Trim()
        $dialogsHandled = ($st.IsTest -and $st.TestLog -match '(?m)^STRETCH dialogs=handled')
        if ($st.IsTest -and $st.Verdict -eq 'PASS' -and -not $dialogsHandled) {
            $st.Verdict = 'FAIL'
            $st.Detail  = ("unacknowledged modal dialog(s) " + $st.Detail).Trim()
        }
    }

    function Write-ItemRecord($item, $st) {
        $rec = [ordered]@{
            type = 'result'; kind = $item.Kind; worker = $WorkerId; run = $item.Run
            test = $item.Rel; suite = $item.Suite; pid = $st.ItemPid
            session = $script:sOrdinal; spos = $st.Spos; served = $null
            verdict = $st.Verdict; outcome = $st.Outcome
            shutdown_s = $st.ShutSecs; test_s = $st.TestSecs
            elapsed_s = [math]::Round($st.Clock.Elapsed.TotalSeconds, 1)
            t_s = [math]::Round($workerClock.Elapsed.TotalSeconds, 1)
            cases_pass = $st.CasesPass; cases_fail = $st.CasesFail
            detail = $st.Detail; dumps = @($st.Dumps); shots = @($st.Shots)
            dialogs = $st.DlgLines.Count; test_exit = $st.TestExit
            log = $st.TestLog; stderr = $st.TestErr
        }
        'RESULT ' + ($rec | ConvertTo-Json -Compress -Depth 4)
    }

    try {
        foreach ($item in $itemQueue) {
            $iterIndex++
            $st = New-ItemState $item
            Close-SessionIfUnusable $st
            if ($null -eq $script:sApp) { Open-Session $item $st }
            elseif ($Cfg.SessionMode -eq 'ReuseClean') { Reset-SessionToBaseline }

            if ($script:sPid -eq 0) {
                Write-UntrackedSessionRecord $item $st
                continue
            }

            $script:sServed++
            $st.Spos = $script:sServed
            if ($st.IsTest -and $null -eq $st.Verdict) { Invoke-TestProcess $item $st $iterIndex }
            Complete-ItemSession $st
            $st.Clock.Stop()
            Resolve-ItemDialogs $st
            Write-ItemRecord $item $st
        }

        # the run is not over until every Excel is accounted for
        if ($script:sApp -or $script:sPid -ne 0) { Close-SessionWithRecord }
    }
    finally {
        # The watchdog is this worker's grandchild, which the parent's cleanup
        # cannot see; it also exits on its own when the worker dies.
        try { Stop-Job $watchdogJob -ErrorAction SilentlyContinue } catch {}
        try { Remove-Job $watchdogJob -Force -ErrorAction SilentlyContinue } catch {}
    }

}

# ---------------------------------------------------------------------------
# Run the workers, streaming results as they arrive: a wedged run has to look
# different from a slow one while it is happening.
# ---------------------------------------------------------------------------
$jobList    = @()
# ALIVE = a reused session survived this test (its close is measured later,
# on the session-close record). Benign for the exit code.
$outcomeSet = @('CLEAN', 'ALIVE', 'CRASH', 'HANG', 'EXITED', 'FAILED')
$verdictSet = @('PASS', 'FAIL', 'SKIP', 'TIMEOUT', 'ERROR')
$outcomeTally = @{}; foreach ($k in $outcomeSet) { $outcomeTally[$k] = 0 }
$verdictTally = @{}; foreach ($k in $verdictSet) { $verdictTally[$k] = 0 }
$shutdowns  = New-Object System.Collections.ArrayList
$seen       = 0
$startedAt  = Get-Date
$lastLine   = Get-Date

function Format-ResultLine($r) {
    # One human line per result; the JSONL holds the whole record.
    $shut = if ($null -eq $r.shutdown_s) { '?' } else { "$($r.shutdown_s)" }
    $runLabel = if ($r.kind -eq 'warmup') { 'warmup' } else { "$($r.run)" }
    $head = if ($r.kind -eq 'session-close') {
        "worker=$($r.worker) session-close session=$($r.session) pid=$($r.pid) served=$($r.served) outcome=$($r.outcome)"
    } elseif ($r.kind -eq 'test') {
        "worker=$($r.worker) test=$($r.test) run=$runLabel pid=$($r.pid) verdict=$($r.verdict) outcome=$($r.outcome)"
    } else {
        "worker=$($r.worker) run=$runLabel pid=$($r.pid) outcome=$($r.outcome)"
    }
    $line = "$head"
    if ($r.kind -ne 'session-close' -and $SessionMode -ne 'Fresh' -and $r.session) {
        # a poisoned-session failure is only attributable when the line says what ran before it
        $line += " session=$($r.session).$($r.spos)"
    }
    if ($r.kind -eq 'session-close' -or $null -ne $r.shutdown_s -or $SessionMode -eq 'Fresh') {
        $line += " shutdown=${shut}s"
    }
    if ($null -ne $r.test_s) { $line += " test_s=$($r.test_s)s" }
    $line += " elapsed=$($r.elapsed_s)s t=$($r.t_s)s"
    if (($r.cases_pass + $r.cases_fail) -gt 0) { $line += " cases=$($r.cases_pass)/$($r.cases_pass + $r.cases_fail)" }
    if ($r.detail) { $line += " $($r.detail)" }
    foreach ($d in @($r.dumps)) { if ($d) { $line += " dump='$d'" } }
    foreach ($s in @($r.shots)) { if ($s) { $line += " shot='$s'" } }
    $line.TrimEnd()
}

# If this script stops for any reason, the finally stops the jobs and kills
# every Excel in the ledger.
try {
    $cursor = 0
    for ($w = 1; $w -le $workerCount; $w++) {
        $slice = @($workItems.GetRange($cursor, $shares[$w - 1]))
        $cursor += $shares[$w - 1]
        $cfg = @{
            WorkerId = $w; Items = $slice; Ledger = $ledgerPath
            CloseTimeout = $CloseTimeoutSeconds; TestTimeout = $TestTimeoutSeconds
            CloseWithX = $CloseWithX.IsPresent; Warmup = $Warmup.IsPresent
            DumpScript = $dumpScript; DumpDir = $dumpRoot; FullDump = $FullDump.IsPresent
            DumpAfter = $DumpAfterSeconds; WorkDir = $workRoot; LogDir = $logRoot
            SessionMode = $SessionMode; CommonScript = $commonScript
        }
        $jobList += Start-Job -ScriptBlock $workerBody -ArgumentList $cfg
    }

    while (@($jobList | Where-Object { $_.State -eq 'Running' -or $_.HasMoreData }).Count -gt 0) {
        foreach ($job in $jobList) {
            # SilentlyContinue: Receive-Job rethrows a child's errors, and with
            # Stop that would abandon every remaining item. They are collected
            # through -ErrorVariable and printed below instead.
            $jobErrs = @()
            foreach ($line in @(Receive-Job $job -ErrorAction SilentlyContinue -ErrorVariable jobErrs)) {
                if ($line -isnot [string]) { continue }
                if ($line -notmatch '^RESULT ') { Write-Output $line; $lastLine = Get-Date; continue }

                $r = $null
                try { $r = ($line.Substring(7) | ConvertFrom-Json) } catch {}
                if ($null -eq $r) { Write-Output $line; continue }

                Write-Output (Format-ResultLine $r)
                # No BOM (5.1's -Encoding utf8 writes one): a .jsonl is read by machines.
                [System.IO.File]::AppendAllText(
                    $resultsPath,
                    $line.Substring(7) + [Environment]::NewLine,
                    (New-Object System.Text.UTF8Encoding($false)))
                $lastLine = Get-Date

                # A warm-up is printed and recorded but never counted, so a known
                # cost stays out of the measurement without being hidden.
                if ($r.kind -eq 'warmup') { continue }
                # a session-close counts into outcomes and shutdowns, not the expected-items check
                if ($r.kind -ne 'session-close') { $seen++ }
                if ($r.outcome) {
                    if ($outcomeTally.ContainsKey($r.outcome)) { $outcomeTally[$r.outcome]++ }
                    else { $outcomeTally[$r.outcome] = 1 }
                }
                if ($r.verdict) {
                    if ($verdictTally.ContainsKey($r.verdict)) { $verdictTally[$r.verdict]++ }
                    else { $verdictTally[$r.verdict] = 1 }
                }
                if ($null -ne $r.shutdown_s) { [void]$shutdowns.Add([double]$r.shutdown_s) }
            }

            # otherwise a worker that died between items leaves no reason anywhere
            foreach ($er in @($jobErrs)) {
                $where = ($er.ScriptStackTrace -replace '\s+', ' ')
                Write-Output ("!!!  worker error: {0}{1}" -f
                              (($er.Exception.Message -replace '\s+', ' ')),
                              $(if ($where) { "  at $where" } else { '' }))
                $lastLine = Get-Date
            }
        }

        # "..." so it can never be mistaken for, or parsed as, a result line
        $quiet = ((Get-Date) - $lastLine).TotalSeconds
        if ($quiet -ge $HeartbeatSeconds) {
            $busy = @($jobList | Where-Object { $_.State -eq 'Running' }).Count
            Write-Output ("...  waiting: {0} worker(s) running, {1} result(s) reported, {2:n0}s since the last one" `
                          -f $busy, $seen, $quiet)
            $lastLine = Get-Date
        }

        Start-Sleep -Milliseconds $ReceivePollMs
    }
}
finally {
    foreach ($job in $jobList) {
        # a Failed job died of a terminating error whose reason is otherwise never shown
        if ($job.State -eq 'Failed') {
            $why = ''
            try { $why = $job.ChildJobs[0].JobStateInfo.Reason.Message } catch {}
            Write-Output ("!!!  worker job {0} FAILED: {1}" -f $job.Id, (($why -replace '\s+', ' ')))
        }
        try { if ($job.State -eq 'Running') { Stop-Job $job -ErrorAction SilentlyContinue } } catch {}
        try { Remove-Job $job -Force -ErrorAction SilentlyContinue } catch {}
    }
    $orphans = Stop-LedgerProcesses $ledgerPath
    if ($orphans -gt 0) {
        Write-Output ("  cleanup: killed {0} Excel process(es) still alive at exit" -f $orphans)
    }
    Remove-Item $logRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Output ''
Write-Output 'StretchXL summary'
Write-Output ("  results       : {0} of {1} expected" -f $seen, $totalItems)
if (-not $floorMode) {
    foreach ($key in ($verdictSet + @($verdictTally.Keys | Where-Object { $verdictSet -notcontains $_ } | Sort-Object))) {
        Write-Output ("  {0,-8}      : {1}" -f $key, $verdictTally[$key])
    }
}
foreach ($key in ($outcomeSet + @($outcomeTally.Keys | Where-Object { $outcomeSet -notcontains $_ } | Sort-Object))) {
    Write-Output ("  {0,-8}      : {1}" -f $key, $outcomeTally[$key])
}
# a rate is only meaningful beside its denominator
if ($seen -gt 0) {
    Write-Output ("  crash rate    : {0:n0}% of {1} result(s)" -f (100 * $outcomeTally.CRASH / $seen), $seen)
    Write-Output ("  hang rate     : {0:n0}% of {1} result(s)" -f (100 * $outcomeTally.HANG  / $seen), $seen)
}

# The shutdown distribution is the result; the hang count is only where the
# threshold fell. Quantiles, because a mean of a bimodal distribution describes
# a shutdown that never happens.
if ($shutdowns.Count -gt 0) {
    $sortedShut = @($shutdowns | Sort-Object)
    $atQuantile = { param($a, $q) $a[[Math]::Min($a.Count - 1, [Math]::Floor($q * $a.Count))] }
    Write-Output ("  shutdown s    : min {0} / p50 {1} / p90 {2} / max {3}   n={4}, over {5}s: {6}" -f `
                  $sortedShut[0], (& $atQuantile $sortedShut 0.5), (& $atQuantile $sortedShut 0.9),
                  $sortedShut[-1], $sortedShut.Count, $SlowShutdownSeconds,
                  @($sortedShut | Where-Object { $_ -gt $SlowShutdownSeconds }).Count)
}
Write-Output ("  elapsed       : {0:n1} min" -f ((Get-Date) - $startedAt).TotalMinutes)
Write-Output ("  results file  : {0}" -f $resultsPath)
if ($seen -ne $totalItems) {
    Write-Output '  WARNING: fewer results than items -- a worker died without reporting.'
}

# 0 only if every result was benign. EXITED counts as bad ("could not read the
# exit code" is not provably clean), and so does a shortfall: missing is not passing.
$badVerdicts = 0
foreach ($k in @('FAIL', 'TIMEOUT', 'ERROR', 'SETUP-FAILED')) {
    if ($verdictTally.ContainsKey($k)) { $badVerdicts += $verdictTally[$k] }
}
$badOutcomes = 0
foreach ($k in @('CRASH', 'HANG', 'EXITED', 'FAILED')) {
    if ($outcomeTally.ContainsKey($k)) { $badOutcomes += $outcomeTally[$k] }
}
$exitCode = if ($badVerdicts -gt 0 -or $badOutcomes -gt 0 -or $seen -ne $totalItems) { 1 } else { 0 }
Write-Output ("  exit code     : {0}" -f $exitCode)
exit $exitCode
