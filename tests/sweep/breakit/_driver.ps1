# A deliberate attempt to break XLL tracing, one hostile input at a time.
#
# Where the xll suite asserts the tracer is RIGHT about ordinary calls, this
# fails loudly on the unambiguous wrongs -- a changed result, an entry without
# an exit, a span whose halves disagree, an exit stamped before its entry --
# and on decoded arguments whose answer is knowable by hand. The rest go to the
# captured log: asserting a guess would pin the guess.

function Invoke-FuzzCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')
    $c = $Case

    try {
        $sx = Connect-TestExcel
        Set-XRaySessionDefaults $sx
        $run = Invoke-XRayFormulaTrace $sx 'Fuzz' @($c.F)
        $baseline = $run.Baseline['A1']
        $now = $run.Now['A1']
        $rows = $run.Rows

        $problems = @()
        # THE UNAMBIGUOUS WRONGS.
        if ($now -ne $baseline) { $problems += "tracing changed the result: '$baseline' -> '$now'" }

        $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') })
        $exits   = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL') })
        $entrySpans = @{}; foreach ($e in $entries) { $entrySpans[$e.span] = ($entrySpans[$e.span] + 1) }
        $exitSpans  = @{}; foreach ($x in $exits)   { $exitSpans[$x.span]  = ($exitSpans[$x.span] + 1) }
        $orphanE = @($entrySpans.Keys | Where-Object { -not $exitSpans.ContainsKey($_) })
        $orphanX = @($exitSpans.Keys  | Where-Object { -not $entrySpans.ContainsKey($_) })
        $dupE    = @($entrySpans.Keys | Where-Object { $entrySpans[$_] -gt 1 })
        if ($orphanE.Count) { $problems += "$($orphanE.Count) entry record(s) with no exit -- reads as a hang" }
        if ($orphanX.Count) { $problems += "$($orphanX.Count) exit record(s) with no entry" }
        if ($dupE.Count)    { $problems += "$($dupE.Count) span id(s) reused by more than one entry" }

        # WHO CALLED IT. Every fuzz case drives its call from a FORMULA in
        # A1, so any entry whose caller is not its cell is misattribution,
        # not fuzz noise -- hostile ARGUMENTS must not corrupt the CALLER.
        $problems += Test-RowInvariants $rows
        # Calls, when the case knows them: every XLL call, in order, and no others.
        if ($c.ContainsKey('Calls')) { $problems += Test-ExpectedTrace $rows $c.Calls 'XLL' }
        # The first traced call is one the formula writes; a built-in around it (SUM, IF) is not traced.
        if ($entries.Count -and $c.F -notmatch ('(?i)\b' + [regex]::Escape([string]$entries[0].function) + '\(')) {
            $problems += "first entry is '$($entries[0].function)', which the formula '$($c.F)' does not call"
        }
        foreach ($e in $entries) {
            if ($e.caller -ne 'cell') { $problems += "entry $($e.function): caller '$($e.caller)', expected 'cell'" }
        }

        foreach ($g in ($rows | Where-Object { $_.span } | Group-Object span)) {
            $fns = @($g.Group | Select-Object -ExpandProperty function -Unique)
            if ($fns.Count -gt 1) { $problems += "span $($g.Name) names more than one function: $($fns -join ', ')" }
            $en = $g.Group | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') } | Select-Object -First 1
            $ex = $g.Group | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL') }  | Select-Object -First 1
            if ($en -and $ex -and [long]$ex.qpc -lt [long]$en.qpc) { $problems += "span $($g.Name): exit qpc precedes entry qpc" }
        }

        # NESTING, when the case declares it. An add-in that re-enters Excel
        # via xlUDF puts one hook genuinely on the stack while another fires;
        # cell-level nesting like =TxB(TxB(1,2),3) does NOT -- Excel evaluates
        # the inner call, finishes it, then the outer. The `depth` and `parent`
        # columns state the nesting, and the order of the rows states it
        # independently: both are checked, against each other.
        if ($c.ContainsKey('NestDepth')) {
            $gotDepth = Get-MaxNestDepth $rows
            if ($gotDepth -ne $c.NestDepth) {
                $problems += ("nesting depth {0} from the entry/exit interleaving, expected {1}" -f $gotDepth, $c.NestDepth)
            }
            # The columns must agree: the deepest XLL entry says so itself, and every
            # nested entry names, as its parent, an entry one level up.
            $colDepth = [int](@($entries | ForEach-Object { [int]$_.depth }) | Measure-Object -Maximum).Maximum
            if ($colDepth -ne $c.NestDepth) { $problems += "depth column reaches $colDepth, expected $($c.NestDepth)" }
            $bySpan = @{}; foreach ($e in $entries) { $bySpan[[string]$e.span] = $e }
            foreach ($e in $entries) {
                if ([int]$e.depth -le 1) { continue }
                $p = $bySpan[[string]$e.parent]
                if (-not $p) { $problems += "$($e.function) at depth $($e.depth) names parent $($e.parent), which is no entry in this trace" }
                elseif ([int]$p.depth -ne [int]$e.depth - 1) { $problems += "$($e.function) at depth $($e.depth) has parent $($p.function) at depth $($p.depth)" }
            }
            Write-Output ("nesting: depth {0} (expected {1})" -f $gotDepth, $c.NestDepth)
        }

        # No trace rows is two different facts. If the CELL shows an Excel
        # error, Excel refused the call upstream and the decoder was never
        # reached -- for hostile input that is survival, and worth knowing
        # (measured: 300 chars into a 255-max byte count is refused by Excel
        # itself, #VALUE!, no call; 400 wide chars are legal and traced).
        # No rows AND no error means the call vanished, which is a failure.
        $refusedUpstream = ($entries.Count -eq 0 -and $now -like '#*')
        if ($entries.Count -eq 0 -and -not $refusedUpstream) {
            $problems += "nothing was traced and the cell shows no error ('$now') -- the call vanished"
        }

        # Only cases whose answer is derivable by hand state these:
        #   Value  what A1 must show;  Args  substrings the first entry's args must contain
        if ($c.Value -and $now -ne $c.Value) {
            $problems += "cell is '$now', expected '$($c.Value)' -- the add-in's answer is a pure function of the input"
        }
        if ($c.Args) {
            if ($entries.Count -eq 0) {
                $problems += "no entry row, so none of the expected arguments could be checked"
            }
            else {
                foreach ($a in $c.Args) {
                    if (-not ([string]$entries[0].args).Contains([string]$a)) {
                        $problems += "args missing '$a' (got '$($entries[0].args)')"
                    }
                }
            }
        }
        #   ArgsExact  the whole of the first entry's args, where a substring cannot tell a cut value from a full one
        if ($c.ArgsExact -and $entries.Count -gt 0 -and [string]$entries[0].args -cne [string]$c.ArgsExact) {
            $problems += "args were '$($entries[0].args)', expected exactly '$($c.ArgsExact)'"
        }
        if ($c.ArgCount -and $entries.Count -gt 0 -and [int]$entries[0].argcount -ne [int]$c.ArgCount) {
            $problems += "argcount=$($entries[0].argcount), expected $($c.ArgCount)"
        }

        # REPORTED, not asserted: what the decoder produced for this input.
        foreach ($e in ($entries | Select-Object -First 4)) {
            Write-Output ("traced: {0} args=[{1}]" -f $e.function, $e.args)
        }

        if ($problems.Count) {
            Complete-Test -Fail -Detail (($problems -join '; ') + " [trying: $($c.W)]")
        }
        $how = if ($refusedUpstream) { 'refused upstream by Excel' } else { 'decoded' }
        Complete-Test -Pass -Detail ("survived ({0}): {1} (cell='{2}', {3} trace rows)" -f $how, $c.W, $now, $rows.Count)
    }
    catch {
        Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
    }
}
