# Stress driver: a real workbook per case (event wiring, buttons, class modules, sheet code,
# cross-workbook dependencies, .xlam add-ins), because in some cases the open itself is the
# trigger. Each .test.ps1 is self-contained: it carries its own case data and hands it here.
#
# The cases are unreasonable on purpose: 500-deep recursion, 60-parameter signatures, a Change
# handler that calls Calculate and re-enters the interpreter mid-frame, two modules exporting
# the same procedure name, a procedure body past the p-code walker's ceiling.
#
# Every case is also checked for the invariants that must hold no matter what: frames opened ==
# frames closed, no hook faults, the circuit breaker still closed, and no procedure lost to a
# full table. A case's own Expect adds only what is specific to it. Excel's lifecycle,
# deadlines, the dialog watchdog and crash attribution are StretchXL's.

# A1 FROM ROW AND COLUMN NUMBERS. UsedRange.Value2 arrives as a block with no
# addresses on it, and its top-left is wherever the used range starts -- so the
# offsets have to be turned back into the addresses a case will ask for.
function Get-A1Address([int]$Row, [int]$Col) {
    $s = ''
    $n = [int]$Col
    while ($n -gt 0) {
        $n--
        # [int] ON BOTH: [Math]::Floor returns a Double, so without it `$n % 26`
        # is a Double and the char cast fails -- for every column past Z, which
        # is exactly the range a quick test of A and B does not reach.
        $s = [char]([int][char]'A' + [int]($n % 26)) + $s
        $n = [int][Math]::Floor($n / 26)
    }
    return "$s$Row"
}

function Get-VbaEntryNames($Totals) {
    # The function each VBA entry row names, in trace order.
    return @($Totals.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' } | ForEach-Object { $_.function })
}

function Assert-VbaTraced($Totals, [string[]]$Names) {
    # $null when every name has a VBA entry row; otherwise the first missing one and what was seen.
    $seen = @(Get-VbaEntryNames $Totals)
    foreach ($n in $Names) {
        if ($seen -notcontains $n) {
            return "$n was not traced; VBA entries: [$((@($seen | Select-Object -Unique)) -join ',')]"
        }
    }
    return $null
}

function Read-CaseCounters($App, [string]$BookLeaf, [string]$Macro) {
    # "Name=N;Name=N" from the case's own VBA. Quoted, like every other Application.Run
    # here: a case's leaf carries hyphens, which Application.Run will not take bare.
    $out = @{}
    foreach ($pair in (([string]$App.Run(("'{0}'!{1}" -f $BookLeaf, $Macro))) -split ';')) {
        if ($pair -match '^\s*(\w+)\s*=\s*(-?\d+)\s*$') { $out[$Matches[1]] = [int]$Matches[2] }
    }
    return $out
}

function Invoke-StressCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')

    $c = $Case

    function New-CaseWorkbook($xl, $spec, $path) {
        $wb = $xl.Workbooks.Add()
        try { $wb.EnableAutoRecover = $false } catch {}
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'S1'
        if ($spec.Modules) {
            foreach ($k in $spec.Modules.Keys) {
                $m = $wb.VBProject.VBComponents.Add(1); $m.Name = $k
                $m.CodeModule.AddFromString($spec.Modules[$k])
            }
        }
        if ($spec.Classes) {
            foreach ($k in $spec.Classes.Keys) {
                $m = $wb.VBProject.VBComponents.Add(2); $m.Name = $k
                $m.CodeModule.AddFromString($spec.Classes[$k])
            }
        }
        if ($spec.SheetCode) {
            # By CODE NAME, not tab name: event handlers live on the component.
            $comp = $wb.VBProject.VBComponents.Item($ws.CodeName)
            $comp.CodeModule.AddFromString($spec.SheetCode)
        }
        if ($spec.ThisWbCode) {
            $wb.VBProject.VBComponents.Item('ThisWorkbook').CodeModule.AddFromString($spec.ThisWbCode)
        }
        if ($spec.Shapes) {
            foreach ($s in $spec.Shapes) {
                $b = $ws.Buttons().Add(20 + 160 * $s.Index, 10, 150, 30)
                $b.Name = $s.Name; $b.Caption = $s.Name; $b.OnAction = $s.OnAction
            }
        }
        if ($spec.ActiveX) {
            $ws.Activate()
            $inserted = $false
            foreach ($ct in @('Forms.CheckBox.1', 'Forms.CommandButton.1')) {
                try {
                    $o = $ws.OLEObjects().Add($ct, $null, $false, $false, $null, $null, $null, 30, 60, 140, 24)
                    $o.Name = $spec.ActiveX
                    $inserted = $true
                    break
                } catch {}
            }
            if (-not $inserted) { throw 'SKIP: this Excel will not insert an ActiveX control' }
        }
        if ($spec.Cells) {
            foreach ($addr in $spec.Cells.Keys) { $ws.Range($addr).Formula = $spec.Cells[$addr] }
        }
        if ($spec.FillFormula) {
            $ws.Range($spec.FillFormula.Range).Formula = $spec.FillFormula.Formula
        }
        Remove-Item $path -Force -ErrorAction SilentlyContinue
        if ($spec.IsAddin) {
            # 55 = xlOpenXMLAddIn: what makes its Public procedures visible to
            # every other workbook unqualified.
            $wb.IsAddin = $true
            $wb.SaveAs($path, 55)
        } else {
            $wb.SaveAs($path, 52)
        }
        $wb.Close($false)
        return $path
    }

    try {
        $sx = Connect-TestExcel
        $app = $sx.App
        Set-XRaySessionDefaults $sx
        $paths = Get-XRayPaths $sx.ProcId

        # Trust must cover VBProject access, or every build below throws.
        try { $null = $app.Workbooks.Item(1).VBProject }
        catch { Complete-Test -Skip -Detail 'VBA project access is not trusted on this machine' }

        # Per-invocation work dir (session pid + this test process's pid): in a reused session
        # an earlier invocation of this case may still hold its files open, such as a loaded
        # .xlam dependency, and SaveAs over the same path fails. Unique paths keep the build
        # writable...
        $caseDirWork = Join-Path $sx.WorkDir ("stress_{0}_{1}" -f $sx.ProcId, $PID)
        New-Item -ItemType Directory -Force $caseDirWork | Out-Null
        $bookPath = Join-Path $caseDirWork ("{0}.xlsm" -f $c.Name)
        $bookLeaf = Split-Path $bookPath -Leaf

        # ...and closing OUR OWN earlier leftovers by leaf keeps the open
        # legal (Excel refuses two open workbooks with one name). Only leaves
        # THIS case owns are touched: the rest of the session's dirt is the
        # realism reuse mode exists for, and stays.
        $ownLeaves = @($bookLeaf)
        if ($c.Deps) {
            foreach ($d in $c.Deps) {
                $ownLeaves += ("{0}.{1}" -f $d.Name, $(if ($d.IsAddin) { 'xlam' } else { 'xlsm' }))
            }
        }
        foreach ($leaf in $ownLeaves) {
            try {
                $prev = $app.Workbooks.Item($leaf)
                $prev.Saved = $true
                $prev.Close($false)
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($prev)
            } catch {}
        }

        [void](New-CaseWorkbook $app $c $bookPath)

        # Dependencies first: they must be OPEN before the case workbook, or
        # its formulas resolve to #REF! and the trace tests nothing.
        $depLeaves = @()
        if ($c.Deps) {
            foreach ($d in $c.Deps) {
                $dp = Join-Path $caseDirWork ("{0}.{1}" -f $d.Name, $(if ($d.IsAddin) { 'xlam' } else { 'xlsm' }))
                [void](New-CaseWorkbook $app $d $dp)
                $depLeaves += (Split-Path $dp -Leaf)
                $dw = $app.Workbooks.Open($dp, 0)
                try { $dw.EnableAutoRecover = $false } catch {}
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($dw) } catch {}
            }
        }

        # Workbook_Open must fire while ARMED, so for that trigger the open
        # happens after arming; everything else opens first.
        $wb = $null
        $openWhileArmed = ($c.Trigger.Kind -eq 'Open')
        if (-not $openWhileArmed) {
            $wb = $app.Workbooks.Open($bookPath, 0)
            try { $wb.EnableAutoRecover = $false } catch {}
        }

        # COUNTED TWICE, AND THE DIFFERENCE IS THIS RUN. Opening the workbook already
        # calculates it, so a volatile UDF has run before arming; this reading is taken
        # while nothing is armed, so it adds no rows to the trace it will be compared with.
        $countersBefore = @{}
        if ($c.Counters -and $wb) {
            try { $countersBefore = Read-CaseCounters $app $bookLeaf $c.Counters }
            catch { Complete-Test -Fail -Detail "the Counters macro '$($c.Counters)' failed before arming: $($_.Exception.Message)" }
        }

        $dlgBefore = @(Get-SessionDialogs).Count
        $mark = Get-LogLength $paths.Log
        $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
        if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
        $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
        if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

        # THE TRIGGER.
        $ran = $true; $err = ''
        try {
            $k = $c.Trigger.Kind
            if ($k -eq 'Open') {
                $wb = $app.Workbooks.Open($bookPath, 0)
                try { $wb.EnableAutoRecover = $false } catch {}
            }
            elseif ($k -eq 'Run') {
                $q = "'{0}'!{1}" -f $bookLeaf, $c.Trigger.Name
                if ($c.Trigger.Args) { $app.Run($q, $c.Trigger.Args[0]) | Out-Null }
                else { $app.Run($q) | Out-Null }
            }
            elseif ($k -eq 'Calc') {
                try { $wb.Worksheets('S1').UsedRange.Dirty() } catch {}
                Invoke-XRayRecalc $app 'Rebuild'
            }
            elseif ($k -eq 'Change') { $wb.Worksheets('S1').Range($c.Trigger.Cell).Value2 = $c.Trigger.Value }
            elseif ($k -eq 'Select') { $wb.Worksheets('S1').Activate(); $wb.Worksheets('S1').Range($c.Trigger.Cell).Select() | Out-Null }
            elseif ($k -eq 'ActiveX') { $wb.Worksheets('S1').OLEObjects($c.Trigger.Name).Object.Value = $true }
            elseif ($k -eq 'Button') {
                # Excel's own macro dispatch -- where a hit-tested click ends
                # up; the UI hit-test itself is not exercised.
                $q = "'{0}'!{1}" -f $bookLeaf, $wb.Worksheets('S1').Buttons($c.Trigger.Name).OnAction
                $app.Run($q) | Out-Null
            }
            else { throw "unknown trigger kind '$k'" }
        } catch { $ran = $false; $err = $_.Exception.Message }
        # every trigger is synchronous; a recalc it started may still be settling
        [void](Wait-XRayCalcDone $app)

        $mark2 = Get-LogLength $paths.Log
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        $totLine = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2 90
        $rows = Read-TraceRows $sx.ProcId

        if (-not $ran -and -not $c.Trigger.MayRaise) {
            Complete-Test -Fail -Detail "trigger raised: $err"
        }
        if (-not $totLine) { Complete-Test -Fail -Detail 'no totals line after disarm' }
        $t = ConvertFrom-XRayTotals $totLine
        $t | Add-Member -NotePropertyName rows -NotePropertyValue $rows -Force

        # THE SHEET, SO AN ORACLE DOES NOT HAVE TO RIDE IN THE TRACE.
        #
        # Expect gets the sheet as well as the rows, so a case can compare VBA's own
        # account without calling a traced procedure that would clutter the trace.
        #
        # Read AFTER disarm, so nothing here can appear in the trace it is being
        # compared against. `Value2`, deliberately: it is what the tracer's own
        # Range description reads, so the two are the same question.
        # One COM call, not one per cell: most cases never read the sheet.
        $cells = @{}
        try {
            $used = $wb.Worksheets('S1').UsedRange
            $vals = $used.Value2
            $r0 = [int]$used.Row; $c0 = [int]$used.Column
            if ($vals -is [Array]) {
                for ($i = 1; $i -le $vals.GetLength(0); $i++) {
                    for ($j = 1; $j -le $vals.GetLength(1); $j++) {
                        $v = $vals[$i, $j]
                        if ($null -eq $v) { continue }
                        $cells[(Get-A1Address ($r0 + $i - 1) ($c0 + $j - 1))] = $v
                    }
                }
            }
            elseif ($null -ne $vals) { $cells[(Get-A1Address $r0 $c0)] = $vals }
        } catch { }
        $t | Add-Member -NotePropertyName cells -NotePropertyValue $cells -Force

        # HOW MANY CALLS VBA ITSELF COUNTED. A case whose call count is Excel's to decide --
        # a recalculation, an event cascade -- counts them in its own VBA and returns them
        # from this macro as "Name=N;Name=N". Run AFTER disarm, so reading them traces nothing,
        # and independent of the tracer: the trace is compared against VBA's own tally.
        $counters = @{}
        if ($c.Counters) {
            $after = @{}
            try { $after = Read-CaseCounters $app $bookLeaf $c.Counters }
            catch { Complete-Test -Fail -Detail "the Counters macro '$($c.Counters)' failed: $($_.Exception.Message)" }
            if (-not $after.Count) { Complete-Test -Fail -Detail "the Counters macro '$($c.Counters)' returned nothing to parse" }
            foreach ($k in $after.Keys) {
                $was = if ($countersBefore.ContainsKey($k)) { $countersBefore[$k] } else { 0 }
                $counters[$k] = $after[$k] - $was
            }
        }
        $t | Add-Member -NotePropertyName counters -NotePropertyValue $counters -Force
        $t | Add-Member -NotePropertyName book  -NotePropertyValue (Split-Path $bookPath -Leaf) -Force

        # WHO CALLED THE FIRST FRAME. The trigger kind determines it: Calc
        # reaches a UDF from its cell; every other kind dispatches a macro
        # or event handler with no caller on a sheet -- kind 'none'.
        # ActiveX stays value-unasserted (unmeasured); Trigger.FirstCaller overrides.
        # The caller invariants are asserted on every row regardless.
        $callerProbs = @(Test-RowInvariants $rows)
        # Every case runs VBA, so a trace with no VBA entry lost it, however the totals read.
        if (-not @($rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' }).Count) {
            Complete-Test -Fail -Detail "no VBA entry row in the trace (statements=$($t.statements))"
        }
        # Calls, when the case knows them: every VBA call in the trace, in order, and no others.
        if ($c.ContainsKey('Calls')) { $callerProbs += Test-ExpectedTrace $rows $c.Calls 'VBA' -AnyOrder:([bool]$c.CallsAnyOrder) }
        # The one depth-capped marker a run past the shadow stack writes, when the case expects it.
        if ($c.ContainsKey('DepthCappedUnder')) {
            $marks = @($rows | Where-Object { $_.kind -eq 'depth-capped' })
            $under = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })[[int]$c.DepthCappedUnder]
            if ($marks.Count -ne 1) { $callerProbs += "depth-capped rows: $($marks.Count), expected 1" }
            elseif (-not $under -or [string]$marks[0].parent -ne [string]$under.span) {
                $callerProbs += "the depth-capped row names parent $($marks[0].parent), expected the deepest recorded call's span $(if ($under) { $under.span })"
            }
        }
        $expectFirst = $c.Trigger.FirstCaller
        if (-not $expectFirst) {
            $expectFirst = switch ($c.Trigger.Kind) {
                'Calc'    { 'cell' }
                'ActiveX' { $null }
                default   { 'none' }
            }
        }
        $firstEntry = @($rows | Where-Object { $_.kind -eq 'entry' } | Select-Object -First 1)
        if ($expectFirst -and $firstEntry.Count -eq 1 -and $firstEntry[0].caller -ne $expectFirst) {
            $callerProbs += "first frame ($($firstEntry[0].function)): caller '$($firstEntry[0].caller)', expected '$expectFirst' for a $($c.Trigger.Kind) trigger"
        }
        if ($callerProbs.Count) { Complete-Test -Fail -Detail ($callerProbs -join '; ') }

        # DIALOGS. The manager's watchdog dismissed and recorded them; the
        # case says whether one was EXPECTED, and this driver keeps the
        # judgement by declaring dialogs handled either way it decides.
        $dlgSeen = @(Get-SessionDialogs).Count - $dlgBefore
        if ($dlgSeen -gt 0 -and -not $c.Trigger.ExpectDialog) {
            $lastDlg = @(Get-SessionDialogs) | Select-Object -Last 1
            Complete-Test -Fail -Detail "Excel raised $dlgSeen modal dialog(s): $lastDlg"
        }
        if ($dlgSeen -eq 0 -and $c.Trigger.ExpectDialog) {
            Complete-Test -Fail -Detail 'expected a modal dialog, none appeared'
        }
        if ($dlgSeen -gt 0) { Write-DialogsHandled }
        # The per-case dialog count rides on the totals object: several Expect
        # blocks assert on it ($t.dialogs -lt 1 => fail).
        $t | Add-Member -NotePropertyName dialogs -NotePropertyValue $dlgSeen -Force

        $why = $null
        if ($t.statements -eq 0 -and $err) {
            # A trigger allowed to raise still has to have RUN: zero statements
            # means the VBA never executed, and the exception says why.
            $why = "trigger executed nothing (statements=0); the call reported: $err"
        }
        if ($t.hookFaults -gt 0) { $why = "$($t.hookFaults) hook fault(s): the tracer faulted inside a VBA thread" }
        elseif ($totLine -match 'BREAKER=OPEN') { $why = 'circuit breaker opened' }
        elseif ($t.framesOpened -ne $t.framesClosed) { $why = "frame leak: opened $($t.framesOpened), closed $($t.framesClosed)" }
        elseif ($t.tableFull -gt 0) { $why = "$($t.tableFull) procedures could not be recorded (table full)" }

        # DID EXCEL STILL COMPUTE THE RIGHT ANSWER? Numeric compare with a
        # tight relative tolerance: formatting differences say nothing about
        # the tracer, a clobbered xmm register is wrong in the first figures.
        if (-not $why -and $c.VerifyCells -and $wb) {
            try {
                $ws2 = $wb.Worksheets('S1')
                foreach ($addr in $c.VerifyCells.Keys) {
                    $want = $c.VerifyCells[$addr]
                    $got  = $ws2.Range($addr).Value2
                    $wd = 0.0; $gd = 0.0
                    # invariant culture: a comma-decimal locale would read "1.5" as 15
                    $inv = [Globalization.CultureInfo]::InvariantCulture
                    $flt = [Globalization.NumberStyles]::Float
                    $wIsNum = [double]::TryParse([string]$want, $flt, $inv, [ref]$wd)
                    $gIsNum = [double]::TryParse([string]$got, $flt, $inv, [ref]$gd)
                    if ($wIsNum -and $gIsNum) {
                        $scale = [Math]::Max([Math]::Abs($wd), 1.0)
                        if ([Math]::Abs($gd - $wd) / $scale -gt 1e-9) {
                            $why = "WRONG RESULT under tracing: $addr is [$got], expected [$want]"; break
                        }
                    } elseif ("$got" -ne "$want") {
                        $why = "WRONG RESULT under tracing: $addr is [$got], expected [$want]"; break
                    }
                }
            } catch { $why = "could not read verification cells: $($_.Exception.Message)" }
        }

        if (-not $why -and $c.Expect) { $why = & $c.Expect $t }

        # Close the case's workbooks clean inside the session -- discard is
        # the close's meaning, and the worker's Saved sweep is the net, not
        # the plan.
        try {
            $leaves = @($bookLeaf) + $depLeaves
            foreach ($w in @($app.Workbooks)) {
                if ($leaves -contains $w.Name) { $w.Saved = $true; $w.Close($false) }
                try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($w) } catch {}
            }
        } catch {}

        if ($why) { Complete-Test -Fail -Detail $why }
        Complete-Test -Pass -Detail ("stmts={0} procs={1} depth={2} frames={3}/{4} rows={5}" -f `
            $t.statements, $t.procedures, $t.maxDepth, $t.framesOpened, $t.framesClosed, $rows.Count)
    }
    catch {
        $msg = $_.Exception.Message
        if ($msg -like '*SKIP:*') {
            Complete-Test -Skip -Detail ($msg -replace '^.*SKIP:\s*', '')
        }
        Complete-Test -Fail -Detail ("exception: " + $msg)
    }
}
