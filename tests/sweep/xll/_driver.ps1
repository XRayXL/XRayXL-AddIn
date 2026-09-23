# The XLL regression driver: one formula in A1, and what the trace says about
# that call -- the decoded arguments and return value -- against answers
# derivable by hand. Each numbered case file carries its own data; several
# share a function, so the file is the case's identity.
#
#   Value    what A1 must hold (proves tracing did not break the call)
#   Args     every argument, in order, each matched whole: `a1:I=7` does not match `a1:I=70000`
#   Ret      the exit row's ret, matched whole
#   TypeText, ArgCount, RetType   optional: when present, typetext and argcount on the entry
#            and rettype on the exit must match exactly
#   Why      what the case defends, for the failure message
#
# An expected value ending in `{` is a PREFIX: the case pins an array's shape, not its elements.

function Invoke-XllCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
    . (Join-Path $PSScriptRoot '..\_xray_common.ps1')
    $c = $Case

    try {
        $sx = Connect-TestExcel
        Set-XRaySessionDefaults $sx
        $run = Invoke-XRayFormulaTrace $sx 'XllCase' @($c.Formula)
        $baseline = $run.Baseline['A1']
        $now = $run.Now['A1']
        $rows = $run.Rows

        # ---- the value: unchanged by tracing, and the right answer --------
        if ($now -ne $baseline) {
            Complete-Test -Fail -Detail "tracing CHANGED the result: '$baseline' -> '$now' (defends: $($c.Why))"
        }
        if ($now -ne $c.Value) {
            Complete-Test -Fail -Detail "cell is '$now', expected '$($c.Value)' (defends: $($c.Why))"
        }

        # ---- the trace: matched by the CELL it reports --------------------
        # ONE CALL: the formula is calculated once while armed, so a second entry is a
        # duplicate row, not a second call.
        $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -and (Get-CallerCell $_) -eq 'A1' })
        if ($entries.Count -eq 0) {
            Complete-Test -Fail -Detail "no entry record naming cell A1 for $($c.Fn) (defends: $($c.Why))"
        }
        if ($entries.Count -gt 1) {
            Complete-Test -Fail -Detail "$($entries.Count) entry records name cell A1, expected 1 (spans $(@($entries.span) -join ','))"
        }
        $e = $entries[0]
        if ($e.function -ne $c.Fn) {
            Complete-Test -Fail -Detail "A1 attributed to '$($e.function)', expected '$($c.Fn)'"
        }
        $x = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL') -and $_.span -eq $e.span })
        if ($x.Count -ne 1) {
            Complete-Test -Fail -Detail "$($x.Count) exit records share span $($e.span), expected 1 -- none reads as a hang"
        }
        $x = $x[0]

        # Whole values; one ending in `{` is a prefix.
        function Test-Whole([string]$Got, [string]$Want) {
            if ($Want.EndsWith('{')) { return $Got.StartsWith($Want, [StringComparison]::Ordinal) }
            return $Got -ceq $Want
        }
        $problems = @()
        $gotArgs = @(if ($e.args) { [regex]::Split([string]$e.args, ' (?=a\d+:)') })
        $wantArgs = @($c.Args)
        if ($gotArgs.Count -ne $wantArgs.Count) {
            $problems += "args has $($gotArgs.Count) argument(s), expected $($wantArgs.Count) (got '$($e.args)')"
        }
        else {
            for ($i = 0; $i -lt $wantArgs.Count; $i++) {
                if (-not (Test-Whole $gotArgs[$i] $wantArgs[$i])) { $problems += "argument $($i + 1) is '$($gotArgs[$i])', expected '$($wantArgs[$i])'" }
            }
        }
        if ($c.ContainsKey('TypeText') -and [string]$e.typetext -cne $c.TypeText) { $problems += "typetext '$($e.typetext)', expected '$($c.TypeText)'" }
        if ($c.ContainsKey('ArgCount') -and [string]$e.argcount -ne $c.ArgCount) { $problems += "argcount '$($e.argcount)', expected '$($c.ArgCount)'" }
        if ($c.ContainsKey('RetType') -and [string]$x.rettype -cne $c.RetType) { $problems += "rettype '$($x.rettype)', expected '$($c.RetType)'" }
        if (-not (Test-Whole ([string]$x.ret) ([string]$c.Ret))) { $problems += "ret is '$($x.ret)', expected '$($c.Ret)'" }
        if (-not (Get-CallerSheet $e)) { $problems += 'no sheet' }
        # every case is a formula in A1, so 'cell' is the only true caller
        if ($e.caller -ne 'cell') { $problems += "caller is '$($e.caller)', expected 'cell' for a formula call" }
        $problems += Test-RowInvariants $rows
        if ($problems.Count) {
            Complete-Test -Fail -Detail (($problems -join '; ') + " (defends: $($c.Why))")
        }

        Complete-Test -Pass -Detail ("{0} args-ok ret-ok cell=A1" -f $c.Fn)
    }
    catch {
        Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
    }
}
