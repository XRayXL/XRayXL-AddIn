# Integration: every xll case on one sheet, in one armed calculation, so adjacent rows exercise
# span allocation and cell attribution against each other. Keyed by cell, not by function name,
# since two cases share a function. The case count is whatever tests/sweep/xll/cases holds.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# The per-case files are the single source of truth: each returns early under $StretchCollectOnly,
# and the numbered filenames give a stable cell order (00-* -> A1, 01-* -> A2, ...).
$allCases = @()
$slugs = @()
$StretchCollectOnly = $true
foreach ($caseFile in (Get-ChildItem (Join-Path $PSScriptRoot 'cases') -Filter '*.test.ps1' | Sort-Object Name)) {
    $case = $null
    . $caseFile.FullName
    if ($case) { $allCases += $case; $slugs += ($caseFile.Name -replace '\.test\.ps1$', '') }
}
$StretchCollectOnly = $false

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    # Every case's formula recalculates while armed, in one pass.
    $run = Invoke-XRayFormulaTrace $sx 'XllAll' @($allCases | ForEach-Object { $_.Formula })
    $baseline = $run.Baseline
    $nowText = $run.Now
    $rows = $run.Rows

    $failCount = 0
    for ($i = 0; $i -lt $allCases.Count; $i++) {
        $c = $allCases[$i]
        $addr = "A$($i + 1)"
        $slug = $slugs[$i]      # the case's own file name, so a report names the file to open
        $why = $null

        if ($nowText[$addr] -ne $baseline[$addr]) {
            $why = "tracing changed the result: '$($baseline[$addr])' -> '$($nowText[$addr])'"
        }
        elseif ($nowText[$addr] -ne $c.Value) {
            $why = "cell is '$($nowText[$addr])', expected '$($c.Value)'"
        }
        else {
            $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') -and (Get-CallerCell $_) -eq $addr })
            if ($entries.Count -eq 0) { $why = "no entry record naming cell $addr" }
            else {
                $e = $entries[0]
                if ($e.function -ne $c.Fn) { $why = "attributed to '$($e.function)', expected '$($c.Fn)'" }
                else {
                    $exitRow = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'XLL') -and $_.span -eq $e.span })
                    if ($exitRow.Count -eq 0) { $why = "no exit sharing span $($e.span)" }
                    else {
                        $problems = @()
                        foreach ($a in $c.Args) { if (-not ([string]$e.args).Contains([string]$a)) { $problems += "args missing '$a'" } }
                        if (-not ([string]$exitRow[0].ret).Contains([string]$c.Ret)) { $problems += "ret missing '$($c.Ret)'" }
                        if ($c.ContainsKey('TypeText') -and [string]$e.typetext -cne $c.TypeText) { $problems += "typetext '$($e.typetext)', expected '$($c.TypeText)'" }
                        if ($c.ContainsKey('ArgCount') -and [string]$e.argcount -ne $c.ArgCount) { $problems += "argcount '$($e.argcount)', expected '$($c.ArgCount)'" }
                        if ($c.ContainsKey('RetType') -and [string]$exitRow[0].rettype -cne $c.RetType) { $problems += "rettype '$($exitRow[0].rettype)', expected '$($c.RetType)'" }
                        if ($problems.Count) { $why = $problems -join '; ' }
                    }
                }
            }
        }

        if ($why) { $failCount++; Write-TestCase -Name $slug -Fail -Detail $why }
        else { Write-TestCase -Name $slug -Pass }
    }

    if ($failCount -eq 0) {
        Complete-Test -Pass -Detail ("all {0} cases in one armed pass, {1} trace rows" -f $allCases.Count, $rows.Count)
    }
    else {
        Complete-Test -Fail -Detail "$failCount of $($allCases.Count) case(s) failed in the burst"
    }
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
