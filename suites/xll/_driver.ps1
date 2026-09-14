# The XLL regression driver: one formula in A1, and what the trace says about
# that call -- the decoded arguments and return value -- against answers
# derivable by hand. Each numbered case file carries its own data; several
# share a function, so the file is the case's identity.
#
#   Value    what A1 must hold (proves tracing did not break the call)
#   Args     substrings that must appear in the entry row's args
#   Ret      substring that must appear in the exit row's ret
#   TypeText, ArgCount, RetType   optional: when present, typetext and argcount on the entry
#            and rettype on the exit must match exactly
#   Why      what the case defends, for the failure message
#
# Substrings, so the row format can grow without rewriting every case, yet
# specific enough to catch a value read from the wrong register.

function Invoke-XllCase($Case) {
    . (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
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
        $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -and (Get-CallerCell $_) -eq 'A1' })
        if ($entries.Count -eq 0) {
            Complete-Test -Fail -Detail "no entry record naming cell A1 for $($c.Fn) (defends: $($c.Why))"
        }
        $e = $entries[0]
        if ($e.function -ne $c.Fn) {
            Complete-Test -Fail -Detail "A1 attributed to '$($e.function)', expected '$($c.Fn)'"
        }
        $x = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL') -and $_.span -eq $e.span })
        if ($x.Count -eq 0) {
            Complete-Test -Fail -Detail "entry has no exit sharing span $($e.span) -- that reads as a hang"
        }
        $x = $x[0]

        $problems = @()
        foreach ($a in $c.Args) {
            if (-not ([string]$e.args).Contains([string]$a)) { $problems += "args missing '$a' (got '$($e.args)')" }
        }
        if ($c.ContainsKey('TypeText') -and [string]$e.typetext -cne $c.TypeText) { $problems += "typetext '$($e.typetext)', expected '$($c.TypeText)'" }
        if ($c.ContainsKey('ArgCount') -and [string]$e.argcount -ne $c.ArgCount) { $problems += "argcount '$($e.argcount)', expected '$($c.ArgCount)'" }
        if ($c.ContainsKey('RetType') -and [string]$x.rettype -cne $c.RetType) { $problems += "rettype '$($x.rettype)', expected '$($c.RetType)'" }
        if (-not ([string]$x.ret).Contains([string]$c.Ret)) { $problems += "ret missing '$($c.Ret)' (got '$($x.ret)')" }
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
        Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
    }
}
