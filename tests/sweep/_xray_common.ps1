# Shared plumbing for talking to the XRayXL XLL in a StretchXL session; the underscore keeps it
# out of test discovery. Product knowledge lives here so StretchXL stays product-agnostic.

function Clear-XRayStaleTraces([int]$ExcelPid) {
    # Pids are reused and the trace file is created lazily, so a test that traces nothing would
    # read a stale same-pid file. Moved aside, not deleted: under Reuse it is an earlier test's.
    if ($ExcelPid -le 0) { return }
    $dir = Join-Path (Get-XRayRoot) 'TraceFiles'
    $stale = @(foreach ($ext in 'csv', 'jsonl') {
        Get-ChildItem (Join-Path $dir ("XRayXL_Trace_*_{0}.{1}" -f $ExcelPid, $ext)) -ErrorAction SilentlyContinue })
    if ($stale.Count -eq 0) { return }
    $earlier = Join-Path $dir 'earlier'
    New-Item -ItemType Directory -Force $earlier | Out-Null
    foreach ($f in $stale) { Move-Item -LiteralPath $f.FullName -Destination $earlier -Force -ErrorAction SilentlyContinue }
}

# Excel's Application events, as issue #13 lists them; the tests' own copy, not read back from the add-in.
$script:AppEvents = @(
    'SheetChange', 'SheetCalculate', 'AfterCalculate', 'SheetTableUpdate', 'WorkbookModelChange',
    'SheetSelectionChange', 'SheetActivate', 'SheetDeactivate', 'WorkbookActivate', 'WorkbookDeactivate',
    'SheetBeforeDoubleClick', 'SheetBeforeRightClick', 'SheetFollowHyperlink',
    'NewWorkbook', 'WorkbookOpen', 'WorkbookBeforeClose', 'WorkbookBeforeSave', 'WorkbookAfterSave',
    'WorkbookBeforePrint', 'WorkbookNewSheet', 'WorkbookNewChart', 'SheetBeforeDelete',
    'WorkbookAddinInstall', 'WorkbookAddinUninstall', 'SheetLensGalleryRenderComplete',
    'WindowActivate', 'WindowDeactivate', 'WindowResize',
    'SheetPivotTableUpdate', 'SheetPivotTableAfterValueChange', 'SheetPivotTableBeforeAllocateChanges',
    'SheetPivotTableBeforeCommitChanges', 'SheetPivotTableBeforeDiscardChanges',
    'WorkbookPivotTableOpenConnection', 'WorkbookPivotTableCloseConnection',
    'WorkbookRowsetComplete', 'WorkbookBeforeXmlImport', 'WorkbookAfterXmlImport',
    'WorkbookBeforeXmlExport', 'WorkbookAfterXmlExport', 'WorkbookSync',
    'ProtectedViewWindowOpen', 'ProtectedViewWindowActivate', 'ProtectedViewWindowDeactivate',
    'ProtectedViewWindowBeforeEdit', 'ProtectedViewWindowBeforeClose', 'ProtectedViewWindowResize',
    'WorkbookBeforeRemoteChange', 'WorkbookAfterRemoteChange', 'RemoteSheetChange',
    'RemoteWorkbookNewSheet', 'RemoteWorkbookNewChart', 'RemoteSheetBeforeDelete', 'RemoteSheetPivotTableUpdate'
)
$script:CalcEvents = @('SheetChange', 'SheetCalculate', 'AfterCalculate', 'SheetTableUpdate',
    'SheetPivotTableUpdate', 'SheetPivotTableAfterValueChange', 'WorkbookModelChange',
    'WorkbookAfterRemoteChange', 'RemoteSheetChange')
$script:SelectionEvents = @('SheetSelectionChange', 'SheetActivate', 'SheetDeactivate', 'WorkbookActivate',
    'WorkbookDeactivate', 'SheetBeforeDoubleClick', 'SheetBeforeRightClick', 'SheetFollowHyperlink',
    'WindowActivate', 'WindowDeactivate', 'ProtectedViewWindowActivate', 'ProtectedViewWindowDeactivate')

function Set-XRayEvents($Sx, [string[]]$On) {
    # Records exactly these events, one setter call each. Returns the echoes that did not agree.
    $bad = @()
    foreach ($name in $script:AppEvents) {
        $want = $On -contains $name
        $echo = [string]$Sx.App.Run('XRayXL_SetTraceParam', 'EVENTS', $name, $want)
        if ($echo -notlike "EVENTS $name=$(if ($want) { 'TRUE' } else { 'FALSE' })*") { $bad += $echo }
    }
    return $bad
}

function Set-XRaySessionDefaults($Sx) {
    # Disarm a session an earlier test left armed: every setter below refuses while armed, so
    # under Reuse one failed test would fail every later one. No add-in throws: not armed.
    $armedNow = $false
    try { $armedNow = [bool]$Sx.App.Run('XRayXL_IsArmed') } catch {}
    if ($armedNow) {
        [void](Invoke-XRayDisarm $Sx)
        $stillArmed = $true
        try { $stillArmed = [bool]$Sx.App.Run('XRayXL_IsArmed') } catch {}
        if ($stillArmed) { throw 'an earlier test left this session armed, and XRayXL_Disarm did not clear it' }
    }
    Clear-XRayStaleTraces $Sx.ProcId
    $app = $Sx.App
    $app.DisplayAlerts = $false
    # a prior test in a reused session can leave EnableEvents False
    try { $app.EnableEvents = $true } catch {}
    # a crashing run otherwise leaves recovery prompts for whoever next opens Excel; drivers
    # re-apply it per workbook, since each reopen makes a new workbook object
    try { $app.AutoRecover.Enabled = $false } catch {}
    # DisplayAlerts does not cover the external-link prompt.
    try { $app.AskToUpdateLinks = $false } catch {}
    # Trace settings are process state a reused session keeps, so each is reset to the default
    # every test starts from; a test wanting otherwise sets its own after this, before arming.
    try { [void]$app.Run('XRayXL_SetTraceParam', 'XLL', 'DEPTH', 'ALL') } catch {}
    try { [void]$app.Run('XRayXL_SetTraceParam', 'VBA', 'DEPTH', 'ALL') } catch {}
    try { [void]$app.Run('XRayXL_SetTraceParam', [Type]::Missing, 'ARGS', $true) } catch {}
    try { [void]$app.Run('XRayXL_SetTraceParam', [Type]::Missing, 'RETVAL', $true) } catch {}
    # the suites run the shipped ring, not an unbuffered 0
    try { [void]$app.Run('XRayXL_SetTraceParam', 'BUFFERSIZE', 64) } catch {}
    try { [void]$app.Run('XRayXL_SetTraceParam', 'BUFFERWHENFULL', 'PAUSE') } catch {}
    # every reader here but the JSONL test's reads CSV
    try { [void]$app.Run('XRayXL_SetTraceParam', 'FORMAT', 'CSV') } catch {}
    # tests rely on this reset instead of restoring what they changed; OBJECTS and BREAKPOINTS
    # are VBA's alone, and refused without a Source
    try { [void]$app.Run('XRayXL_SetTraceParam', 'VBA', 'OBJECTS', $true) } catch {}
    try { [void]$app.Run('XRayXL_SetTraceParam', 'VBA', 'BREAKPOINTS', $false) } catch {}
    $level = if ($env:XRAYXL_LOGLEVEL) { $env:XRAYXL_LOGLEVEL } else { 'INFO' }
    try { [void]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', $level) } catch {}
    try { [void](Set-XRayEvents $Sx $script:CalcEvents) } catch {}
    # Excel's own default; the timeline driver forces a thread count after this.
    try { $app.MultiThreadedCalculation.Enabled = $true; $app.MultiThreadedCalculation.ThreadMode = 0 } catch {}
}

function Get-XRayRoot {
    # The manager sets XRAYXL_OUTPUT_DIR per session so each Excel's files sit beside the run.
    if ($env:XRAYXL_OUTPUT_DIR) { return $env:XRAYXL_OUTPUT_DIR }
    return (Join-Path $env:TEMP 'XRayXL')
}

function Get-XRayPaths([int]$ExcelPid) {
    # No trace path here: its name carries an unpredictable rising id, so Get-XRayTraceCsv
    # resolves it at read time.
    @{
        Log    = Join-Path (Get-XRayRoot) ("Logs\XRayXL_{0}.log" -f $ExcelPid)
    }
}

function Get-XRayTraceCsv([int]$ExcelPid) {
    # The file is created on the first traced row, so a session that traced nothing gets a
    # non-existent path, which Read-TraceFile reads as no rows.
    $dir = Join-Path (Get-XRayRoot) 'TraceFiles'
    $f = @(Get-ChildItem (Join-Path $dir ("XRayXL_Trace_*_{0}.csv" -f $ExcelPid)) -ErrorAction SilentlyContinue |
           Sort-Object LastWriteTime) | Select-Object -Last 1
    if ($f) { return $f.FullName }
    return (Join-Path $dir ("XRayXL_Trace_none_{0}.csv" -f $ExcelPid))
}

function Get-LogLength([string]$LogPath) {
    (@(Get-Content $LogPath -ErrorAction SilentlyContinue)).Count
}

function Wait-LogLine([string]$LogPath, [string]$Pattern, [int]$Mark, [int]$Seconds = 60) {
    # polls rather than sleeping a guessed interval: arming can take several seconds
    $dl = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $dl) {
        $all = @(Get-Content $LogPath -ErrorAction SilentlyContinue)
        if ($all.Count -gt $Mark) {
            $hit = $all[$Mark..($all.Count - 1)] | Where-Object { $_ -match $Pattern } | Select-Object -Last 1
            if ($hit) { return $hit }
        }
        Start-Sleep -Milliseconds 300
    }
    return $null
}

function Wait-XRayCondition([scriptblock]$Until, [int]$Seconds = 30, [int]$PollMs = 100) {
    # a probe that throws counts as not yet
    $deadline = (Get-Date).AddSeconds($Seconds)
    while ($true) {
        try { if (& $Until) { return $true } } catch {}
        if ((Get-Date) -ge $deadline) { return $false }
        Start-Sleep -Milliseconds $PollMs
    }
}

function Wait-XRayCalcDone($App, [int]$Seconds = 60) {
    # CalculationState 0 is xlDone; Excel rejects COM calls while busy, which reads as not done.
    Wait-XRayCondition { $App.CalculationState -eq 0 } $Seconds
}

function Invoke-XRayRecalc($App, [ValidateSet('Full', 'Rebuild')][string]$How = 'Full', [int]$Seconds = 60) {
    # Returns nothing, so it can sit inside helpers that return an object.
    if ($How -eq 'Rebuild') { $App.CalculateFullRebuild() } else { $App.CalculateFull() }
    [void](Wait-XRayCalcDone $App $Seconds)
}

function Wait-XRayTraceClosed([int]$ExcelPid, [int]$Seconds = 10) {
    # Disarm closes the trace file; an open that will not share with a writer proves it has.
    $path = Get-XRayTraceCsv $ExcelPid
    if (-not (Test-Path $path)) { return $true }
    Wait-XRayCondition { $fs = [IO.File]::Open($path, 'Open', 'Read', 'Read'); $fs.Dispose(); $true } $Seconds
}

function Invoke-XRayCommand($Sx, [string]$Command) {
    # The fault sentinel differs: Arm returns 0 on a contained fault, while Disarm returns its
    # dropped-row count, so its 0 is a clean disarm and -1 the fault.
    try {
        $rc = $Sx.App.Run($Command)
        $faulted = if ($Command -eq 'XRayXL_Disarm') { $rc -eq -1 } else { $rc -eq 0 }
        if ($faulted) { return "$Command faulted inside the XLL (contained -- see XRayXL_crash log)" }
        return 'pressed'
    } catch { return "Application.Run('$Command') failed: $($_.Exception.Message)" }
}

function Invoke-XRayDisarm($Sx) {
    # Returns the drop count, or -1 on a contained fault; the trace file carries no drop
    # marker, so this is how a caller learns the trace is complete.
    $drops = -1
    try { $drops = [int]$Sx.App.Run('XRayXL_Disarm') } catch {}
    [void](Wait-XRayTraceClosed $Sx.ProcId)
    return $drops
}

function Stop-XRayTrace($Sx) {
    # Returns '' or a problem string, so drops, untyped values and p-code walk warnings fail
    # every test rather than a later assertion blaming the tracer. Tests wanting drops use
    # Invoke-XRayDisarm.
    $log  = (Get-XRayPaths $Sx.ProcId).Log
    $mark = Get-LogLength $log
    $drops = Invoke-XRayDisarm $Sx
    $after = @(Get-Content $log -ErrorAction SilentlyContinue | Select-Object -Skip $mark)
    $untyped = @($after | Select-String ' WARNING - VBA (args|returns):' | ForEach-Object { $_.Line -replace '^.*? WARNING - ', '' })
    if ($untyped.Count) {
        return ('UNTYPED VALUE: the disarm report says ' + ($untyped -join ' | '))
    }
    $walks = @($after | Select-String ' WARNING - VBA p-code:' | ForEach-Object { $_.Line -replace '^.*? WARNING - ', '' })
    if ($walks.Count) {
        return ('P-CODE WALK: the disarm report says ' + ($walks -join ' | '))
    }
    if ($drops -lt 0) {
        return 'XRayXL_Disarm faulted inside the XLL (contained -- see the XRayXL crash log)'
    }
    if ($drops -gt 0) {
        return ("TRACE INCOMPLETE: the output buffer dropped $drops row(s) during this run -- " +
                'nothing below can be asserted against a trace with holes in it ' +
                '(raise BUFFERSIZE, or the run is emitting faster than the drain)')
    }
    # A clean disarm ends the trace with its disarm row; without one the file reads as a crash.
    # The newest file in either format: a JSON Lines session writes no .csv.
    $dir = Join-Path (Get-XRayRoot) 'TraceFiles'
    $trace = @(Get-ChildItem (Join-Path $dir ("XRayXL_Trace_*_{0}.*" -f $Sx.ProcId)) -ErrorAction SilentlyContinue |
               Where-Object { $_.Extension -in '.csv', '.jsonl' } | Sort-Object LastWriteTime) |
             Select-Object -Last 1 | ForEach-Object { $_.FullName }
    if ($trace -and (Test-Path $trace)) {
        $last = Get-Content $trace -Tail 1
        $isDisarm = if ($trace -like '*.jsonl') {
            $r = $last | ConvertFrom-Json
            $r.kind -eq 'event' -and $r.source -eq 'XRayXL' -and $r.function -eq 'disarm'
        } else { $last -match '^\d+,\d+,event,XRayXL,0,0,0,\d+,\d+,,disarm,' }
        if (-not $isDisarm) {
            return "TRACE UNFINISHED: a clean disarm left no disarm row last in $trace"
        }
    }
    return ''
}

# Shared Check for multi-case tests. Dot-sourcing resets the failure count.
$script:XRayCheckFails = 0

function Reset-XRayChecks { $script:XRayCheckFails = 0 }

function Get-XRayCheckFailures { return $script:XRayCheckFails }

function Check([string]$Name, [bool]$Ok, [string]$Detail) {
    Write-TestCase $Name -Pass:$Ok -Fail:(-not $Ok) -Detail $Detail
    if (-not $Ok) { $script:XRayCheckFails++ }
}

function Write-XRayObservation([string]$Name, [string]$Detail) {
    # A measurement with no right answer to hold it to: logged, never counted as a case.
    Write-Output ("observed {0}: {1}" -f $Name, (($Detail -replace '\s+', ' ').Trim()))
}

# CVErr codes as Range.Value2 carries them, by the name Excel shows.
$script:XRayErrorNames = @{
    (-2146826288) = '#NULL!'; (-2146826281) = '#DIV/0!'; (-2146826273) = '#VALUE!'
    (-2146826265) = '#REF!';  (-2146826259) = '#NAME?';  (-2146826252) = '#NUM!'
    (-2146826246) = '#N/A';   (-2146826245) = '#GETTING_DATA'; (-2146826243) = '#SPILL!'
    (-2146826242) = '#CONNECT!'; (-2146826241) = '#BLOCKED!'; (-2146826240) = '#UNKNOWN!'
    (-2146826239) = '#FIELD!'; (-2146826238) = '#CALC!'
}

function ConvertTo-XRayCellText($Value) {
    # Range.Value2 as text that reads the same in every locale: numbers to 15
    # significant digits with a '.' separator, TRUE/FALSE, errors by name.
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { if ($Value) { return 'TRUE' } else { return 'FALSE' } }
    if ($Value -is [int]) {
        if ($script:XRayErrorNames.ContainsKey($Value)) { return $script:XRayErrorNames[$Value] }
        return "#ERROR($Value)"
    }
    if ($Value -is [double]) {
        if ($Value -eq 0) { return '0' }       # -0 and 0 are one value to Excel
        return $Value.ToString('G15', [Globalization.CultureInfo]::InvariantCulture)
    }
    return [string]$Value
}

function Get-XRayCellText($Range) { return (ConvertTo-XRayCellText $Range.Value2) }

function ConvertFrom-XRayTraceSummary($Value) {
    # XRayXL_GetTraceSummary's grid (Source, Module, Function, Calls) from either
    # shape Application.Run hands back: 2-D, or flattened row-major.
    $grid = @(ConvertTo-XRayGrid $Value)[0]
    $cells = @()
    if ($grid.Count -gt 0 -and $grid[0] -is [array]) { foreach ($row in $grid) { $cells += @($row) } }
    else { $cells = @($grid) }
    $all = @()
    for ($i = 0; $i + 3 -lt $cells.Count; $i += 4) {
        $all += [pscustomobject]@{
            Source = [string]$cells[$i]; Module = [string]$cells[$i + 1]
            Function = [string]$cells[$i + 2]; Calls = [string]$cells[$i + 3]
        }
    }
    # The header is dropped by name, so a missing one cannot swallow a data row.
    $header = $null
    if ($all.Count -and $all[0].Function -eq 'Function') { $header = $all[0]; $all = @($all | Select-Object -Skip 1) }
    return [pscustomobject]@{ Header = $header; Rows = $all }
}

function Get-XRaySummaryCalls($Summary, [string]$Function) {
    # The call count for one function, 0 when the summary does not list it.
    $row = @($Summary.Rows | Where-Object { $_.Function -eq $Function }) | Select-Object -First 1
    if ($row) { return [double]$row.Calls }
    return [double]0
}

function Close-OwnLeftover($App, [string]$Leaf) {
    # Tests leave their books open, and names are keyed by pid, so under Reuse SaveAs would meet
    # our own earlier book; close that one, by leaf, and touch nothing else.
    if (-not $App -or -not $Leaf) { return }
    try {
        $prev = $App.Workbooks.Item($Leaf)
        try { $prev.Saved = $true } catch {}
        $prev.Close($false)
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($prev) } catch {}
    } catch {}
}

function New-XRayMacroBook {
    # Builds, saves and reopens a workbook (formulas must run as loaded).
    #   Stem        leaf is "<Stem>_<pid>.<Format>"; a reused session may meet the same test twice
    #   Components  in order, @{ Kind = 1|2|3|'Sheet'; Name; Code } -- a class must precede its users;
    #               Code may be an array, added piece by piece
    #   Cells       address -> formula, written after the code
    #   Prepare     scriptblock given the sheet, for anything Cells cannot express
    #   Leaf        a fixed leaf, kept unique by a "<Stem>_<pid>" subdirectory instead
    #   Format      xlsm, or xlsx for a book with no code
    # Books stay open together, keyed by Stem. Returns nothing (use Get-XRayMacroBook):
    # an assigned result would swallow Complete-Test's verdict line.
    param(
        $Sx,
        [string]$Stem,
        [array]$Components = @(),
        [hashtable]$Cells = @{},
        [scriptblock]$Prepare = $null,
        [string]$SheetName = 'S1',
        [string]$Leaf = '',
        [ValidateSet('xlsm', 'xlsx')][string]$Format = 'xlsm'
    )
    $app = $Sx.App
    if ($Leaf) {
        $dir = Join-Path $Sx.WorkDir ("{0}_{1}" -f $Stem, $Sx.ProcId)
        New-Item -ItemType Directory -Force $dir | Out-Null
        $bookPath = Join-Path $dir $Leaf
        if ($Leaf -match '\.xlsx$') { $Format = 'xlsx' } else { $Format = 'xlsm' }
    }
    else {
        $bookPath = Join-Path $Sx.WorkDir ("{0}_{1}.{2}" -f $Stem, $Sx.ProcId, $Format)
    }
    $bookLeaf = Split-Path $bookPath -Leaf
    Close-OwnLeftover $app $bookLeaf          # a reused session may still hold OUR previous one
    Remove-Item $bookPath -ErrorAction SilentlyContinue

    $booksRef = $app.Workbooks
    $wb = $booksRef.Add()
    try { $wb.EnableAutoRecover = $false } catch {}
    $ws = $wb.Worksheets.Item(1); $ws.Name = $SheetName

    # only a book with code needs the VBA project, so an untrusted machine still runs the rest
    $proj = $null
    if ($Components.Count) {
        try { $proj = $wb.VBProject }
        catch { Complete-Test -Skip -Detail 'VBA project access is not trusted on this machine' }
    }

    $added = @()
    foreach ($comp in $Components) {
        # the sheet module already exists; reach it by code name, not tab name
        if ($comp.Kind -eq 'Sheet') {
            foreach ($piece in @($comp.Code)) { $proj.VBComponents.Item($ws.CodeName).CodeModule.AddFromString($piece) }
            continue
        }
        $made = $null
        try { $made = $proj.VBComponents.Add([int]$comp.Kind); $made.Name = $comp.Name }
        catch {
            # a UserForm needs the MSForms designer: a missing prerequisite, so SKIP
            if ([int]$comp.Kind -eq 3) {
                Complete-Test -Skip -Detail "this Excel will not add a UserForm component: $($_.Exception.Message)"
            }
            Complete-Test -Fail -Detail "could not add VBA component '$($comp.Name)': $($_.Exception.Message)"
        }
        foreach ($piece in @($comp.Code)) {
            try { $made.CodeModule.AddFromString($piece) }
            catch {
                # some modules are generated; show the source that failed to compile
                Write-Output $piece
                Complete-Test -Fail -Detail "generated source for '$($comp.Name)' did not compile: $($_.Exception.Message)"
            }
        }
        $added += $made
    }

    foreach ($addr in $Cells.Keys) { $ws.Range($addr).Formula = $Cells[$addr] }
    if ($Prepare) { & $Prepare $ws }

    # Excel refuses SaveAs past 218 characters with an error that never mentions the path.
    if ($bookPath.Length -gt 218) {
        Complete-Test -Fail -Detail (
            ("workbook path is {0} characters and Excel refuses SaveAs past 218, " +
             "so this is the -OutDir, not the test: shorten it. Path: {1}") -f
            $bookPath.Length, $bookPath)
    }

    $fileFormat = if ($Format -eq 'xlsx') { 51 } else { 52 }
    $wb.SaveAs($bookPath, $fileFormat); $wb.Close($false)
    foreach ($r in ($added + @($proj, $ws, $wb))) {
        if ($null -eq $r) { continue }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($r) } catch {}
    }

    # a saved-and-reopened workbook takes a different calc path than a fresh one
    $wb = $booksRef.Open($bookPath)
    # EnableAutoRecover is per workbook object, and the reopen made a new one
    try { $wb.EnableAutoRecover = $false } catch {}

    if (-not $script:XRayMacroBooks) { $script:XRayMacroBooks = @{} }
    $script:XRayMacroBook = [pscustomobject]@{
        Path  = $bookPath
        Leaf  = $bookLeaf
        Book  = $wb
        Sheet = $wb.Worksheets.Item($SheetName)
    }
    $script:XRayMacroBooks[$Stem] = $script:XRayMacroBook
}

function Get-XRayMacroBook([string]$Stem = '') {
    # The book built last, or the one built under $Stem.
    if ($Stem) { return $script:XRayMacroBooks[$Stem] }
    return $script:XRayMacroBook
}

function Invoke-XRayFormulaTrace {
    # Formulas down column A of a saved-and-reopened book, calculated unarmed and
    # then armed. Returns the leaf, the sheet, each cell's value before and during
    # tracing (keyed by address), and the book's rows from the trace.
    # Problems throw: the caller assigns the result, which would swallow a verdict line.
    param($Sx, [string]$Stem, [string[]]$Formulas)
    $cells = @{}
    for ($i = 0; $i -lt $Formulas.Count; $i++) { $cells["A$($i + 1)"] = $Formulas[$i] }
    New-XRayMacroBook $Sx $Stem -Cells $cells -Format xlsx
    $book = Get-XRayMacroBook $Stem
    $log = (Get-XRayPaths $Sx.ProcId).Log

    # the unarmed pass proves the add-in is sane and gives the comparison its left side
    Invoke-XRayRecalc $Sx.App
    $baseline = @{}
    foreach ($addr in $cells.Keys) { $baseline[$addr] = Get-XRayCellText $book.Sheet.Range($addr) }

    $mark = Get-LogLength $log
    $pressed = Invoke-XRayCommand $Sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { throw "arm: $pressed" }
    $armLine = Wait-LogLine $log 'armed \d+ of|nothing armed|could not' $mark
    if (-not $armLine) { throw 'no arm outcome in the action log -- refusing to assert against an unarmed session' }
    if ($armLine -notmatch 'armed \d+ of') { throw "arming did not succeed: $armLine" }

    Invoke-XRayRecalc $Sx.App
    $now = @{}
    foreach ($addr in $cells.Keys) { $now[$addr] = Get-XRayCellText $book.Sheet.Range($addr) }
    $lossy = Stop-XRayTrace $Sx
    if ($lossy) { throw $lossy }

    return [pscustomobject]@{
        Leaf     = $book.Leaf
        Sheet    = $book.Sheet
        Baseline = $baseline
        Now      = $now
        Rows     = @(Select-BookRows (Read-TraceRows $Sx.ProcId) $book.Leaf)
    }
}

function Invoke-XRayArmedSession {
    # One arming session: apply the settings, arm, run Body, disarm.
    #   Settings  @(Source, Name, Value) triples, in order; a $null Source means both
    #   Body      runs while armed; what it returns comes back as Result
    #   Leaf      when set, Rows holds this book's rows from the session's trace
    #   ArmWait   the log line that says arming finished
    # Returns ArmLine, Result, Totals (the VBA totals line, or $null) and Rows.
    # Body's output is Result, so checks are reported after the session, and problems throw.
    # One setting still needs the outer array: -Settings @(,@('XLL', 'DEPTH', 'TOP')).
    # Locals carry an xr prefix so Body, run in a child scope, still sees the caller's names.
    param(
        $Sx,
        [array]$Settings = @(),
        [scriptblock]$Body = $null,
        [string]$Leaf = '',
        # 'XLL tracing: OFF' is the last line an arm writes when the XLL side is off
        [string]$ArmWait = 'armed \d+ of|nothing armed|could not|XLL tracing: OFF'
    )
    foreach ($xrSetting in $Settings) { [void](Set-XRayTraceParam $Sx $xrSetting[0] $xrSetting[1] $xrSetting[2]) }
    $xrLog = (Get-XRayPaths $Sx.ProcId).Log
    $xrMark = Get-LogLength $xrLog
    $xrPressed = Invoke-XRayCommand $Sx 'XRayXL_Arm'
    if ($xrPressed -ne 'pressed') { throw "arm: $xrPressed" }
    $xrArmLine = Wait-LogLine $xrLog $ArmWait $xrMark
    $xrResult = $null
    if ($Body) { $xrResult = & $Body }
    $xrMark2 = Get-LogLength $xrLog
    $xrLossy = Stop-XRayTrace $Sx
    if ($xrLossy) { throw $xrLossy }
    # the log is written synchronously, and a VBA side that was not armed writes no totals
    $xrTotals = @(Get-Content $xrLog -ErrorAction SilentlyContinue | Select-Object -Skip $xrMark2 |
                  Where-Object { $_ -match 'VBA trace: statements=' }) | Select-Object -Last 1
    $xrRows = @()
    if ($Leaf) { $xrRows = @(Select-BookRows (Read-TraceRows $Sx.ProcId) $Leaf) }
    return [pscustomobject]@{ ArmLine = $xrArmLine; Result = $xrResult; Totals = $xrTotals; Rows = $xrRows }
}

function Set-XRayTraceParam($Sx, $Source, [string]$Name, $Value) {
    # $null Source means both: an omitted XLL argument arrives as xltypeMissing in any position.
    # Callers assert the echo: a setter that records a value and changes nothing looks the same.
    $s = if ($null -eq $Source) { [Type]::Missing } else { $Source }
    try { return [string]$Sx.App.Run('XRayXL_SetTraceParam', $s, $Name, $Value) }
    catch { return "Application.Run('XRayXL_SetTraceParam') failed: $($_.Exception.Message)" }
}

function ConvertTo-XRayGrid($Value) {
    # Excel hands a COM array back 1-based, so the bounds are read: indexing from 0 drops a row.
    if ($null -eq $Value) { return @() }
    if ($Value -isnot [array]) { return @(,@([string]$Value)) }
    # Application.Run flattens an XLL array row-major, losing only the shape; tests assert the
    # shape through a spilled cell.
    if ($Value.Rank -eq 1) { return ,@($Value | ForEach-Object { [string]$_ }) }
    $rows = @()
    for ($r = $Value.GetLowerBound(0); $r -le $Value.GetUpperBound(0); $r++) {
        $row = @()
        for ($c = $Value.GetLowerBound(1); $c -le $Value.GetUpperBound(1); $c++) {
            $row += [string]$Value[$r, $c]
        }
        $rows += ,$row
    }
    # PowerShell unrolls one level on return; a bare $rows would hand back the first row
    return ,$rows
}

function Get-XRayTraceParam($Sx, $Source, $Name) {
    # Scalar when both Source and Name are given, otherwise a 2-D array; $null means not supplied.
    $s = if ($null -eq $Source) { [Type]::Missing } else { $Source }
    $n = if ($null -eq $Name)   { [Type]::Missing } else { $Name }
    try { return $Sx.App.Run('XRayXL_GetTraceParam', $s, $n) }
    catch { return "Application.Run('XRayXL_GetTraceParam') failed: $($_.Exception.Message)" }
}

function ConvertFrom-XRayTotals([string]$Line) {
    # "VBA trace: statements=N exits=N ..." -> object; absent keys read as 0.
    $t = @{}
    foreach ($m in [regex]::Matches($Line, '(\w+)=(\d+)')) { $t[$m.Groups[1].Value] = [int64]$m.Groups[2].Value }
    return [pscustomobject]@{
        statements = [int64]$t['statements']; exits = [int64]$t['exits']
        transitions = [int64]$t['transitions']; procedures = [int64]$t['procedures']
        maxDepth = [int64]$t['maxDepth']; recursions = [int64]$t['recursions']
        faults = [int64]$t['faults']; stackGrowFailures = [int64]$t['stackGrowFailures']
        tableFull = [int64]$t['tableFull']; unmatchedExits = [int64]$t['unmatchedExits']
        named = [int64]$t['named']; unnamed = [int64]$t['unnamed']
        framesClosed = [int64]$t['framesClosed']; framesOpened = [int64]$t['framesOpened']
        hookFaults = [int64]$t['hookFaults']
        ipEntries = [int64]$t['ipEntries']; ipStaleClosed = [int64]$t['ipStaleClosed']
        ipUnavailable = [int64]$t['ipUnavailable']
        ipIntraProc = [int64]$t['ipIntraProc']; ipLateOpen = [int64]$t['ipLateOpen']
        returnsRead = [int64]$t['returnsRead']; returnsDeclined = [int64]$t['returnsDeclined']
        returnsOff = [int64]$t['returnsOff']
        byrefChanged = [int64]$t['byrefChanged']; byrefSame = [int64]$t['byrefSame']
        byrefDeclined = [int64]$t['byrefDeclined']; doEventsChains = [int64]$t['doEventsChains']
    }
}

# The trace-file contract (docs/TraceRowModel.md). The header is the version, so an old reader
# meeting a new format refuses loudly rather than mis-filtering silently.
$script:TraceHeader = 'seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,caller,callerref,argcount,args,ret,rettype,outcome,ticks,tracerticks,trust'
# kind and source are separate columns so a filter on one need not spell out the other
$script:TraceKinds   = @('entry', 'exit', 'event')
$script:TraceSources = @('XLL', 'VBA', 'Excel', 'XRayXL')

# In .NET because a PowerShell character loop is too slow for traces of millions of rows.
# One row is one line: the writer turns CR and LF inside a field into spaces.
if (-not ('XRayCsvShape' -as [type])) {
    Add-Type -TypeDefinition @'
public static class XRayCsvShape
{
    // null when every line has `want` fields; otherwise a description of the
    // FIRST line that does not -- the reader turns that into the throw.
    public static string FirstBadLine(string path, int want)
    {
        long lineNo = 0;
        foreach (string line in System.IO.File.ReadLines(path))
        {
            lineNo++;
            if (lineNo == 1) continue;            // the header, already checked
            if (line.Length == 0) continue;
            int n = 1; bool inQuote = false;
            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (c == '"') inQuote = !inQuote;
                else if (c == ',' && !inQuote) n++;
            }
            if (n != want)
                return "line " + lineNo + " has " + n + " columns, the header names " + want;
        }
        return null;
    }
}
'@
}

function Test-TraceColumnCount([string]$Path, [int]$Want) {
    return [XRayCsvShape]::FirstBadLine((Resolve-Path $Path).ProviderPath, $Want)
}


function Read-TraceFile([string]$Path) {
    # A missing file is an empty result: "never armed" is a fact, not a fault. A present file
    # that violates the contract throws, so every test that reads the trace also checks its
    # format.
    if (-not (Test-Path $Path)) { return @() }

    # CSV is the only format this reader knows; a new one adds a branch here, not a new reader.
    if ($Path -notmatch '\.csv$') {
        throw "trace contract: no reader for '$Path' -- formats are added to docs/TraceRowModel.md and Read-TraceFile in the same change"
    }

    # an optional `breaks` column may follow, under VBA BREAKPOINTS
    $first = Get-Content $Path -TotalCount 1
    $hasBreaks = ($first -ceq "$script:TraceHeader,breaks")
    if ($first -cne $script:TraceHeader -and -not $hasBreaks) {
        throw ("trace contract violated at {0}`n  expected: {1}`n  found:    {2}" -f $Path, $script:TraceHeader, $first)
    }

    # Import-Csv pads a short row with nulls that look like empty fields, so only counting
    # commas catches a row that lost a column.
    $bad = Test-TraceColumnCount $Path ($first.Split(',')).Count
    if ($bad) { throw "trace contract: $bad, in $Path" }

    $rows = @(Import-Csv $Path)   # a parse failure here throws -- loud on purpose
    $prev = [int64]0
    $inSeen = @{}
    foreach ($r in $rows) {
        $seq = [int64]0
        if (-not [int64]::TryParse($r.seq, [ref]$seq)) { throw "trace contract: non-numeric seq '$($r.seq)' in $Path" }
        if ($seq -le $prev) { throw "trace contract: seq $seq after $prev -- not strictly increasing in $Path" }
        $prev = $seq
        # `input` is the producer's emit sequence: a hole is a dropped row, and concurrent
        # producers reorder it, so only uniqueness is checked.
        $in = [int64]0
        if (-not [int64]::TryParse($r.input, [ref]$in) -or $in -lt 1) { throw "trace contract: bad input '$($r.input)' at seq $seq in $Path" }
        if ($inSeen.ContainsKey($in)) { throw "trace contract: input $in used twice (at seq $seq) in $Path" }
        $inSeen[$in] = $true
        if ($script:TraceKinds -notcontains $r.kind) { throw "trace contract: unknown kind '$($r.kind)' at seq $seq in $Path" }
        # loss is reported by Disarm and the log, never as a row
        if ($script:TraceSources -notcontains $r.source) { throw "trace contract: unknown source '$($r.source)' at seq $seq in $Path" }
        $u = [uint64]0
        if (-not [uint64]::TryParse($r.span,   [ref]$u)) { throw "trace contract: non-numeric span '$($r.span)' at seq $seq in $Path" }
        if (-not [uint64]::TryParse($r.thread, [ref]$u)) { throw "trace contract: non-numeric thread '$($r.thread)' at seq $seq in $Path" }
        $q = [int64]0
        if (-not [int64]::TryParse($r.qpc, [ref]$q)) { throw "trace contract: non-numeric qpc '$($r.qpc)' at seq $seq in $Path" }
        # breaks: a count on every VBA exit row, empty on every other row
        if ($hasBreaks) {
            $b = [uint32]0
            if ($r.kind -eq 'exit' -and $r.source -eq 'VBA') {
                if (-not [uint32]::TryParse($r.breaks, [ref]$b)) { throw "trace contract: bad breaks '$($r.breaks)' on a VBA exit at seq $seq in $Path" }
            } elseif ($r.breaks -ne '') { throw "trace contract: breaks '$($r.breaks)' on a $($r.source) $($r.kind) row at seq $seq in $Path" }
        }
    }
    # The session's own rows: arm first, and a disarm, when there is one, last.
    if ($rows.Count) {
        $arm = $rows[0]
        if ($arm.kind -ne 'event' -or $arm.source -ne 'XRayXL' -or $arm.function -ne 'arm') {
            throw "trace contract: the first row is $($arm.source) $($arm.kind) $($arm.function), not the arm row, in $Path"
        }
        $disarms = @($rows | Where-Object { $_.source -eq 'XRayXL' -and $_.function -eq 'disarm' })
        if ($disarms.Count -gt 1) { throw "trace contract: $($disarms.Count) disarm rows in $Path" }
        if ($disarms.Count -eq 1 -and $disarms[0].seq -ne $rows[-1].seq) {
            throw "trace contract: the disarm row (seq $($disarms[0].seq)) is not the last in $Path"
        }
    }
    return $rows
}

function Get-MaxNestDepth($Rows) {
    # XLL nesting from the entry/exit interleaving, independent of the `depth` column, so a
    # nested call written as two sequential ones is caught.
    $spans = @{}
    foreach ($r in $Rows) {
        if (-not $r.span) { continue }
        if ($r.source -ne 'XLL') { continue }
        if ($r.kind -ne 'entry' -and $r.kind -ne 'exit') { continue }
        if (-not $spans.ContainsKey($r.span)) {
            $spans[$r.span] = [pscustomobject]@{ Span = $r.span; Thread = $r.thread; Enter = $null; Exit = $null }
        }
        if ($r.kind -eq 'entry') { $spans[$r.span].Enter = [long]$r.seq }
        else                     { $spans[$r.span].Exit  = [long]$r.seq }
    }
    # an unclosed span is the orphan check's failure; guessing here would report it twice
    $iv = @($spans.Values | Where-Object { $null -ne $_.Enter -and $null -ne $_.Exit })
    $max = 0
    foreach ($a in $iv) {
        $d = 1
        foreach ($b in $iv) {
            if ($b.Span -eq $a.Span) { continue }
            if ($b.Thread -eq $a.Thread -and $b.Enter -lt $a.Enter -and $b.Exit -gt $a.Exit) { $d++ }
        }
        if ($d -gt $max) { $max = $d }
    }
    return $max
}

function Select-BookRows($Rows, [string]$BookLeaf) {
    # Under Reuse, CalculateFull recalculates leftover books whose rows collide with this test's.
    # Scoped by span, so an exit follows its entry; rows tied to no book (a macro, an event) stay.
    # An empty trace arrives as $null, which a pipeline would pass on as one empty row.
    $Rows = @($Rows | Where-Object { $null -ne $_ })
    $mine = @{}; $theirs = @{}
    foreach ($r in $Rows) {
        $sh = Get-CallerSheet $r
        if (-not $sh) { continue }
        if ($sh.Contains($BookLeaf)) { $mine[[string]$r.span] = $true }
        else                         { $theirs[[string]$r.span] = $true }
    }
    @($Rows | Where-Object {
        $sh = Get-CallerSheet $_
        if ($sh)       { return $sh.Contains($BookLeaf) }
        $s = [string]$_.span
        if (-not $s)   { return $true }              # nothing to scope it by
        if ($mine.ContainsKey($s))   { return $true }
        return -not $theirs.ContainsKey($s)
    })
}

function Read-TraceRows([int]$ExcelPid) {
    Read-TraceFile (Get-XRayTraceCsv $ExcelPid)
}

# The kinds `caller` may take, and whether `callerref` must describe it. Closed on purpose: a new
# kind is a decoder change and should fail here loudly.
$script:CallerKinds = @{
    'cell'        = $true     # "[Book1]Sheet1!B2", or a whole CSE range
    'name'        = $true     # a shape's name, or an Auto_* macro's sheet name
    'toolbar'     = $true     # position, then the bar: "2/5" or "2/\"MyBar\""
    'menu'        = $true     # command, menu, bar, submenu
    'editor'      = $true     # the VBA editor started it: Run, F8 or the Immediate window
    'registerid'  = $true     # the DLL called itself
    'none'        = $true     # ref | nil | emptyref | sheetless-B2 | nametoolong
    'unavailable' = $true     # Excel declined; the xlret code
    'array'       = $true     # an array of a size nobody documents
    'unknown'     = $true     # an XLOPER type we have not seen
}

# A `cell` callerref is an external address as Range.Address(,,,True) gives it. Excel quotes the
# `[Book]Sheet` prefix when a name needs it and doubles apostrophes, so both forms are accepted.
$script:ExternalAddress = "^(?:\[[^\]]+\][^!]+|'\[[^\]]+\](?:[^']|'')*')!"

# No quoting logic needed: an address has exactly one `!`.
function Get-CallerCell($Row) {
    if ($Row.caller -ne 'cell') { return '' }
    if ($Row.callerref -match '^(.+)!([^!]+)$') { return $Matches[2] }
    return ''
}

function Get-CallerSheet($Row) {
    if ($Row.caller -ne 'cell') { return '' }
    if ($Row.callerref -match '^(.+)!([^!]+)$') { return $Matches[1] }
    return ''
}

# How an activation ended; an XLL exit can only be `returned`.
$script:Outcomes = @('returned', 'threw', 'unwound', 'handled', 'abandoned', 'unhandled')

# What ended the measurement, not a verdict on it: `exit` and `end` are readings, `backstop` and
# `flush` upper bounds, and `async` means no duration exists yet.
$script:Trusts = @('exit', 'end', 'backstop', 'flush', 'async')

function Test-RowInvariants($Rows) {
    # Invariants that hold on any correct trace (docs/TraceRowModel.md); returns problem strings.
    $problems = @()
    foreach ($r in $Rows) {
        # An event is not a call: no span, no chain, no duration; its parameters are named.
        if ($r.kind -eq 'event') {
            if ($r.source -ne 'Excel' -and $r.source -ne 'XRayXL') { $problems += "seq $($r.seq): event from source '$($r.source)'" }
            if ($r.source -eq 'XRayXL' -and $r.function -ne 'arm' -and $r.function -ne 'disarm') {
                $problems += "seq $($r.seq): XRayXL event '$($r.function)' is not arm or disarm"
            }
            if (-not $r.function) { $problems += "seq $($r.seq): event with no name" }
            foreach ($c in 'span', 'parent', 'depth') {
                if ([string]$r.$c -ne '0') { $problems += "seq $($r.seq): event row's $c is '$($r.$c)', not 0" }
            }
            foreach ($c in 'typetext', 'caller', 'argcount', 'ret', 'rettype', 'outcome', 'ticks', 'tracerticks', 'trust') {
                if ($r.$c) { $problems += "seq $($r.seq): event row carries $c '$($r.$c)'" }
            }
            if ($r.args -and [string]$r.args -notmatch '^[A-Za-z_]\w*:') {
                $problems += "seq $($r.seq): event args '$($r.args)' is not <name>:<type>=<value>"
            }
            continue
        }
        if ($r.kind -eq 'entry') {
            # typetext has no brackets; no parameters is an empty typetext with argcount 0.
            if ([string]$r.typetext -like '(*') { $problems += "seq $($r.seq): typetext '$($r.typetext)' still carries brackets" }
            if ([string]$r.argcount -eq '0' -and $r.typetext) { $problems += "seq $($r.seq): argcount 0 but typetext '$($r.typetext)'" }
            if (-not $r.caller) { $problems += "seq $($r.seq): entry with EMPTY caller"; continue }
            if (-not $script:CallerKinds.ContainsKey($r.caller)) {
                $problems += "seq $($r.seq): caller kind '$($r.caller)' is not one this reader knows"
                continue
            }
            $wants = $script:CallerKinds[$r.caller]
            $has   = [bool]$r.callerref
            if ($wants -and -not $has) { $problems += "seq $($r.seq): caller='$($r.caller)' but callerref is empty" }
            if (-not $wants -and $has) { $problems += "seq $($r.seq): caller='$($r.caller)' must carry no callerref, has '$($r.callerref)'" }
            # anything but an external address has lost its book or sheet
            if ($r.caller -eq 'cell' -and $r.callerref -notmatch $script:ExternalAddress) {
                $problems += "seq $($r.seq): caller='cell' but callerref '$($r.callerref)' is not an external address"
            }
        }
        elseif ($r.kind -eq 'exit') {
            if ($r.caller -or $r.callerref) { $problems += "seq $($r.seq): exit row carries caller/callerref" }
            if ($r.typetext -or $r.argcount) { $problems += "seq $($r.seq): exit row carries typetext/argcount" }
        }
        # One args grammar on both sources: every argument starts a<N>:.
        if ($r.args -and [string]$r.args -notmatch '^a\d+:') {
            $problems += "seq $($r.seq): $($r.source) args '$($r.args)' is not a<N>:<type>=<value>"
        }

        # an XLL exit row exists only because the call returned
        if ($r.kind -eq 'exit') {
            if (-not $r.outcome) { $problems += "seq $($r.seq): $($r.source) exit with EMPTY outcome" }
            elseif ($script:Outcomes -notcontains $r.outcome) {
                $problems += "seq $($r.seq): outcome '$($r.outcome)' is not one this reader knows"
            }
            elseif ($r.source -eq 'XLL' -and $r.outcome -ne 'returned') {
                $problems += "seq $($r.seq): XLL exit reads outcome '$($r.outcome)'; only 'returned' is possible"
            }
        }
        elseif ($r.outcome) {
            $problems += "seq $($r.seq): $($r.source) $($r.kind) row carries outcome '$($r.outcome)'"
        }

        # An async XLL call is the one exit with no ticks, and says so in `trust`. An entry row
        # has no duration, so a value in either is one the writer invented.
        if ($r.kind -eq 'exit') {
            if (-not $r.trust) { $problems += "seq $($r.seq): exit with no trust" }
            elseif ($script:Trusts -notcontains $r.trust) {
                $problems += "seq $($r.seq): trust '$($r.trust)' is not one this reader knows"
            }
            if (-not $r.ticks -and $r.trust -ne 'async') {
                $problems += "seq $($r.seq): exit with no ticks and trust='$($r.trust)' -- only async has no duration"
            }
            if ($r.ticks -and $r.trust -eq 'async') {
                $problems += "seq $($r.seq): trust='async' but ticks='$($r.ticks)' -- the work has not finished"
            }
            # The tracer's share of a duration: present exactly when ticks is, and never more.
            if ([bool]$r.ticks -ne [bool]$r.tracerticks) {
                $problems += "seq $($r.seq): ticks='$($r.ticks)' but tracerticks='$($r.tracerticks)'"
            }
            elseif ($r.ticks -and [uint64]$r.tracerticks -gt [uint64]$r.ticks) {
                $problems += "seq $($r.seq): tracerticks $($r.tracerticks) exceeds ticks $($r.ticks)"
            }
        }
        else {
            if ($r.ticks) { $problems += "seq $($r.seq): $($r.kind) row carries ticks '$($r.ticks)'" }
            if ($r.tracerticks) { $problems += "seq $($r.seq): $($r.kind) row carries tracerticks '$($r.tracerticks)'" }
            if ($r.trust) { $problems += "seq $($r.seq): $($r.kind) row carries trust '$($r.trust)'" }
        }

        # `parent` 0 is the top of a chain, a value; only an empty field is a fault
        if ('' -eq [string]$r.depth)  { $problems += "seq $($r.seq): $($r.source) row with no depth" }
        elseif ([int]$r.depth -lt 1)  { $problems += "seq $($r.seq): depth '$($r.depth)' is below 1" }
        if ('' -eq [string]$r.parent) { $problems += "seq $($r.seq): $($r.source) row with no parent" }
        if ([string]$r.depth -eq '1' -and [string]$r.parent -ne '0') {
            $problems += "seq $($r.seq): depth=1 but parent='$($r.parent)' -- a chain's first frame has no caller"
        }
        if ([string]$r.depth -ne '' -and [int]$r.depth -gt 1 -and [string]$r.parent -eq '0') {
            $problems += "seq $($r.seq): depth=$($r.depth) but parent=0 -- a nested frame names no caller"
        }
    }

    # a row written twice reuses its span
    $spanSeen = @{}
    foreach ($r in $Rows) {
        if (-not $r.span -or ($r.kind -ne 'entry' -and $r.kind -ne 'exit')) { continue }
        $key = "$($r.kind)|$($r.span)"
        if ($spanSeen.ContainsKey($key)) { $problems += "seq $($r.seq): a second $($r.kind) row for span $($r.span)" }
        else { $spanSeen[$key] = $true }
    }

    # A missing exit reads as a hang. Callers that filter keep both rows of a span.
    foreach ($r in $Rows) {
        if (-not $r.span -or $r.kind -eq 'event') { continue }
        if ($r.kind -eq 'entry' -and -not $spanSeen.ContainsKey("exit|$($r.span)")) {
            $problems += "seq $($r.seq): $($r.source) entry $($r.function) (span $($r.span)) has no exit"
        }
        elseif ($r.kind -eq 'exit' -and -not $spanSeen.ContainsKey("entry|$($r.span)")) {
            $problems += "seq $($r.seq): $($r.source) exit $($r.function) (span $($r.span)) has no entry"
        }
    }

    # VBA leaves module empty when it cannot resolve it, by design
    foreach ($r in $Rows) {
        if ($r.kind -ne 'entry' -and $r.kind -ne 'exit') { continue }
        if (-not $r.function) { $problems += "seq $($r.seq): $($r.source) $($r.kind) with no function" }
        if ($r.source -eq 'XLL' -and -not $r.module) { $problems += "seq $($r.seq): XLL $($r.kind) with no module" }
    }

    # A parent is one level up. Depth is per source, so only a same-source parent is compared.
    $entryBySpan = @{}
    foreach ($r in $Rows) { if ($r.kind -eq 'entry' -and $r.span) { $entryBySpan[[string]$r.span] = $r } }
    foreach ($r in $Rows) {
        if ($r.kind -ne 'entry' -or [string]$r.parent -in @('', '0')) { continue }
        $p = $entryBySpan[[string]$r.parent]
        if ($p -and $p.source -eq $r.source -and [int]$p.depth -ne [int]$r.depth - 1) {
            $problems += "seq $($r.seq): $($r.function) at depth $($r.depth) names parent $($p.function) at depth $($p.depth)"
        }
    }

    # ticks is the exit's qpc less its entry's; checked only where the entry survived filtering
    $entryQpc = @{}
    foreach ($r in $Rows) { if ($r.kind -eq 'entry' -and $r.span) { $entryQpc[[string]$r.span] = [string]$r.qpc } }
    foreach ($r in $Rows) {
        if ($r.kind -ne 'exit' -or -not $r.ticks) { continue }
        $ticks = 0L
        if (-not [long]::TryParse([string]$r.ticks, [ref]$ticks) -or $ticks -lt 0) {
            $problems += "seq $($r.seq): ticks '$($r.ticks)' is not a count of 0 or more"
            continue
        }
        $start = $entryQpc[[string]$r.span]
        if ($start -and ([long]$r.qpc - [long]$start) -ne $ticks) {
            $problems += "seq $($r.seq): ticks $ticks but exit qpc - entry qpc = $([long]$r.qpc - [long]$start) (span $($r.span))"
        }
    }
    return $problems
}

# A missing row reads '(no row)', not $null, so it shows in a failure message.

function EntryRowOf($Rows, [string]$Fn) {
    $r = @($Rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
    if ($r.Count) { return $r[0] }
    return $null
}

function ExitRowOf($Rows, [string]$Fn) {
    $r = @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
    if ($r.Count) { return $r[0] }
    return $null
}

# `args` without the storage addresses, which change every run, for tests that compare values.
function Remove-ArgAddress([string]$Text) { $Text -replace '@0x[0-9A-F]+=', '=' }

function ArgsOf($Rows, [string]$Fn) {
    $r = EntryRowOf $Rows $Fn
    if ($null -eq $r) { return '(no row)' }
    return [string]$r.args
}

function SigOf($Rows, [string]$Fn) {
    $r = EntryRowOf $Rows $Fn
    if ($null -eq $r) { return '(no row)' }
    return [string]$r.typetext
}

function RetOf($Rows, [string]$Fn) {
    $r = ExitRowOf $Rows $Fn
    if ($null -eq $r) { return '(no row)' }
    return [string]$r.ret
}

function ExitArgsOf($Rows, [string]$Fn) {
    # ByRef arguments are re-read at exit: these are the out-values
    $r = ExitRowOf $Rows $Fn
    if ($null -eq $r) { return '(no row)' }
    return [string]$r.args
}

function Outcome($Rows, [string]$Fn) {
    # '(no row)' = no exit at all; '(none)' = exit with a blank outcome, a different fault
    $r = ExitRowOf $Rows $Fn
    if ($null -eq $r) { return '(no row)' }
    if ($r.outcome) { return [string]$r.outcome }
    return '(none)'
}

function ArgOf($Row, [int]$N) {
    # a value may contain spaces, colons and '=', so it ends at the next "aN:", not the next space
    $a = [string]$Row.args
    if ($a -match "(?:^|\s)a$N`:[^=\s]*=(.*?)(?=\s+a\d+:[^=\s]*=|`$)") { return $Matches[1] }
    return ''
}

function Test-WholeValue([string]$Got, [string]$Want) {
    # Whole values; an expected value ending in `{` pins an array's shape, so it is a prefix.
    if ($Want.EndsWith('{')) { return $Got.StartsWith($Want, [StringComparison]::Ordinal) }
    return $Got -ceq $Want
}

function Get-CallMismatches($Want, $Entry, $Exit) {
    # The fields of one expected call that the entry and its exit do not match.
    $checks = [ordered]@{
        Function = [string]$Entry.function; Args = [string]$Entry.args; Depth = [string]$Entry.depth
        Caller = [string]$Entry.caller; Cell = (Get-CallerCell $Entry)
        Ret = [string]$Exit.ret; RetType = [string]$Exit.rettype; Outcome = [string]$Exit.outcome
    }
    foreach ($k in $checks.Keys) {
        if ($Want.ContainsKey($k) -and -not (Test-WholeValue $checks[$k] ([string]$Want[$k]))) {
            "$k is '$($checks[$k])', expected '$($Want[$k])'"
        }
    }
}

function Test-ExpectedTrace($Rows, [array]$Expected, [string]$Source, [switch]$AnyOrder) {
    # The calls the test knows it made, and no others. $Expected lists every $Source call in entry
    # order; Parent is the index of the enclosing call, or -1. -AnyOrder is for calls whose order
    # Excel decides, and ignores Parent. The count is checked first: nothing lines up without it.
    $entries = @($Rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq $Source })
    $want = @($Expected)
    if ($entries.Count -ne $want.Count) {
        return @("$Source entries: $($entries.Count), expected $($want.Count) -- got [$(@($entries.function) -join ',')], expected [$(@($want | ForEach-Object { $_.Function }) -join ',')]")
    }
    $exitBySpan = @{}
    foreach ($r in $Rows) { if ($r.kind -eq 'exit' -and $r.source -eq $Source) { $exitBySpan[[string]$r.span] = $r } }
    $problems = @()
    if ($AnyOrder) {
        $used = @{}
        for ($i = 0; $i -lt $want.Count; $i++) {
            $match = $null
            foreach ($e in $entries) {
                if ($used.ContainsKey([string]$e.span)) { continue }
                $x = $exitBySpan[[string]$e.span]
                if ($x -and -not @(Get-CallMismatches $want[$i] $e $x).Count) { $match = $e; break }
            }
            if ($match) { $used[[string]$match.span] = $true }
            else { $problems += "$Source call $($i + 1) ($($want[$i].Function)): no entry matches $(($want[$i].Keys | Sort-Object | ForEach-Object { "$_=$($want[$i][$_])" }) -join ' ')" }
        }
        return $problems
    }
    for ($i = 0; $i -lt $want.Count; $i++) {
        $w = $want[$i]; $e = $entries[$i]; $x = $exitBySpan[[string]$e.span]
        $at = "$Source call $($i + 1) ($($w.Function))"
        if ([string]$e.function -cne $w.Function) { $problems += "${at}: the trace names '$($e.function)'"; continue }
        if (-not $x) { $problems += "${at}: no exit row"; continue }
        foreach ($m in @(Get-CallMismatches $w $e $x)) { $problems += "${at}: $m" }
        if ($w.ContainsKey('Parent')) {
            $wantParent = if ([int]$w.Parent -lt 0) { '0' } else { [string]$entries[[int]$w.Parent].span }
            if ([string]$e.parent -ne $wantParent) { $problems += "${at}: parent is span $($e.parent), expected $wantParent" }
        }
    }
    return $problems
}

function Test-TracedCallsMatchCounters($Totals) {
    # Where Excel decides the call count, VBA counts its own calls: a tally independent of the
    # tracer, where the tracer's own counters would prove only self-consistency.
    if (-not $Totals.counters -or -not $Totals.counters.Count) { return 'the case declared no counters' }
    foreach ($fn in ($Totals.counters.Keys | Sort-Object)) {
        $traced = @($Totals.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq $fn }).Count
        if ($traced -ne $Totals.counters[$fn]) {
            return "${fn}: $traced entry row(s), but VBA counted $($Totals.counters[$fn]) call(s)"
        }
    }
    return $null
}

function Get-FirstEntryByName($Rows, [string]$Source = 'VBA') {
    # first, not last: later calls of the same name carry other arguments
    $map = @{}
    foreach ($r in $Rows) {
        if ($r.kind -ne 'entry') { continue }
        if ($r.source -ne $Source) { continue }
        if (-not $map.ContainsKey($r.function)) { $map[$r.function] = $r }
    }
    return $map
}

function Read-XRayNames([string]$LogPath, [int]$Mark) {
    # the report's resolved names are the only way to tell a name from the right name
    $names = @()
    $all = @(Get-Content $LogPath -ErrorAction SilentlyContinue)
    if ($all.Count -gt $Mark) {
        foreach ($ln in $all[$Mark..($all.Count - 1)]) {
            if ($ln -match '^\s+\[[^\]]+\]\S*?\.(\w+)\s+\d+\s+\d+') { $names += $Matches[1] }
            elseif ($ln -match '^\s+(0x[0-9A-Fa-f]+) \(unnamed\)') { $names += '(unnamed)' }
        }
    }
    return $names
}
