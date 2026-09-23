# Call counts belong to the arming session, not the process.
#
# XRayXL_GetTraceSummary reports a call count per function, from Target::calls. Each arm writes
# its own trace file, created lazily, so a re-arm that traces nothing produces no file and the
# summary must likewise report nothing.
#
# Targets come from a fixed array and a slot is handed straight back on the next arm, so `calls`
# has to be reset at arm. A slot is bound to an index, not a function: if the registration table
# changes between arms, a stale count would be reported against a different function's name.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    function Get-TracedNames {
        @((ConvertFrom-XRayTraceSummary ($app.Run('XRayXL_GetTraceSummary', 'Tx*'))).Rows | ForEach-Object { $_.Function })
    }
    function Get-TxBCalls {
        Get-XRaySummaryCalls (ConvertFrom-XRayTraceSummary ($app.Run('XRayXL_GetTraceSummary', 'Tx*'))) 'TxB'
    }

    # AsLoaded: a formula in a saved workbook is what a user has.
    New-XRayMacroBook $sx 'XllCounts' -Cells @{ 'A1' = '=TxB(2,3)' } -Format xlsx

    $settings = @(@('XLL', 'DEPTH', 'ALL'), @('VBA', 'DEPTH', 'OFF'))

    # ---- session 1: call it several times ---------------------------------
    # Rebuild, not Calculate: a non-volatile UDF served from its last result
    # would not be called again and the count would not move.
    $s1 = Invoke-XRayArmedSession $sx -Settings $settings -Body {
        1..3 | ForEach-Object { Invoke-XRayRecalc $app 'Rebuild' }
        Get-TxBCalls
    }
    Check 'first-arm-armed' ([bool]($s1.ArmLine -match 'armed \d+ of')) "$($s1.ArmLine)"
    $first = [double]$s1.Result
    Check 'first-session-counts-its-calls' ($first -ge 3) "TxB=$first after 3 rebuilds"
    # Session 1 traced, so it wrote a file; a re-arm that traces nothing must not write a newer one.
    $s1file = Get-XRayTraceCsv $sx.ProcId

    # ---- session 2: arm again, read before calculating anything, then one recalc
    # Nothing has been called yet, so the summary must name nothing: a stale
    # count is not zero and would survive the summary's calls == 0 skip.
    $s2 = Invoke-XRayArmedSession $sx -Settings $settings -Body {
        $namedBefore = Get-TracedNames
        $fileBefore = Get-XRayTraceCsv $sx.ProcId
        Invoke-XRayRecalc $app 'Rebuild'
        [pscustomobject]@{ NamedBefore = @($namedBefore); FileBefore = $fileBefore; Calls = (Get-TxBCalls) }
    }
    Check 'second-arm-armed' ([bool]($s2.ArmLine -match 'armed \d+ of')) "$($s2.ArmLine)"
    Check 're-arm-starts-with-nothing-traced' ($s2.Result.NamedBefore -notcontains 'TxB') `
          ("summary named: " + (($s2.Result.NamedBefore | Sort-Object) -join ',') + " -- expected none")

    # The trace file is the second witness: lazy creation means no new file yet.
    Check 'nothing-traced-wrote-no-new-file' ($s2.Result.FileBefore -eq $s1file) `
          "newest after a no-trace re-arm: $(Split-Path $s2.Result.FileBefore -Leaf); session 1's was $(Split-Path $s1file -Leaf)"

    $second = [double]$s2.Result.Calls
    Check 'second-session-counts-its-own-calls' ($second -ge 1) "TxB=$second after 1 rebuild"
    # How many calls a rebuild makes is Excel's business, so this is an inequality:
    # per-session is strictly smaller than three rebuilds, cumulative strictly larger.
    Check 'counts-did-not-carry-over' ($second -lt $first) `
          "session 1 (3 rebuilds) TxB=$first, session 2 (1 rebuild) TxB=$second -- carried over would exceed it"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "counts reset at arm: 3 rebuilds gave $first, a fresh arm gave nothing, then 1 rebuild gave $second"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
