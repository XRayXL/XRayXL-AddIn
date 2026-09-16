# The trace-file contract, proven NEGATIVELY: feed Read-TraceFile doctored
# files and assert it refuses each one the way docs/TraceRowModel.md
# promises. The ~170 product tests prove the reader ACCEPTS what the tracer
# writes; without this file the refusal paths -- the entire point of a strict
# reader -- are the untested code. Needs no arming and never binds the session the
# manager provides; the good-file corpus is real rows, verbatim, from real
# traced sessions.
#
# The header is RESTATED LITERALLY here rather than read from
# $script:TraceHeader, deliberately: this file is an independent witness.
# Edit the constant in _xray_common.ps1 without editing the doc and this
# test, and this test fails -- which is the evolution policy working.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$hdr = 'seq,input,kind,source,span,parent,depth,thread,qpc,module,function,proc,typetext,caller,callerref,argcount,args,ret,rettype,outcome,ticks,trust'
# `input` (col 2) is the producer's emit sequence; here it has a HOLE at 3 --
# input 1,2,4,5 across four rows means one row was dropped, exactly what the
# reader must accept (holes are drops) while still requiring uniqueness.
$good = @(
    '1,1,entry,XLL,74079595921409,,,17248,2759222870812,TracedAddin64.xll,TxB,TxB,"B,B",cell,[XllCase_10020.xlsx]Sheet1!A1,2,a1:B=2 a2:B=3,,,,,'
    '2,2,exit,XLL,74079595921409,,,17248,2759222871416,TracedAddin64.xll,TxB,TxB,,,,,,23,Q,returned,604,exit'
    '3,4,entry,VBA,1,0,1,17212,2767391810634,[err-resume-next-1000.xlsm]M,Go,0x25E154EF3E4,,none,ref,0,,,,,,'
    '4,5,exit,VBA,1,0,1,17212,2767391816836,[err-resume-next-1000.xlsm]M,Go,0x25E154EF3E4,,,,,,,,returned,4541,exit'
)

# This test never binds an Excel, so it has no $sx to take a WorkDir from --
# but it still writes files, and they belong with every other test's output
# rather than loose in the shared work root.
$dir = Get-TestWorkDir -Managed ([bool]$env:STRETCH_TEST_ID)
$script:failed = 0

function New-TraceFile([string]$Name, [string[]]$Lines) {
    $p = Join-Path $dir $Name
    Set-Content -Path $p -Value ($Lines -join "`r`n") -Encoding Ascii
    return $p
}

function Expect-Refusal([string]$Case, [string]$Pattern, [string]$Path) {
    try {
        $null = Read-TraceFile $Path
        Write-TestCase $Case -Fail -Detail 'ACCEPTED what the contract forbids'
        $script:failed++
    } catch {
        if ($_.Exception.Message -match $Pattern) { Write-TestCase $Case -Pass }
        else {
            Write-TestCase $Case -Fail -Detail "refused for the wrong reason: $($_.Exception.Message)"
            $script:failed++
        }
    }
}

# ---- the accept path, on real rows ----------------------------------------
try {
    $rows = @(Read-TraceFile (New-TraceFile 'good.csv' (@($hdr) + $good)))
    if ($rows.Count -eq 4 -and $rows[0].kind -eq 'entry' -and $rows[0].callerref -eq '[XllCase_10020.xlsx]Sheet1!A1' -and
        $rows[3].ticks -eq '4541' -and $rows[3].trust -eq 'exit') { Write-TestCase 'accepts-real-rows' -Pass }
    else { Write-TestCase 'accepts-real-rows' -Fail -Detail "got $($rows.Count) rows"; $script:failed++ }
} catch { Write-TestCase 'accepts-real-rows' -Fail -Detail $_.Exception.Message; $script:failed++ }

# A header-only file is a session that armed and traced nothing: zero rows,
# no complaint.
try {
    $rows = @(Read-TraceFile (New-TraceFile 'empty-session.csv' @($hdr)))
    if ($rows.Count -eq 0) { Write-TestCase 'accepts-empty-session' -Pass }
    else { Write-TestCase 'accepts-empty-session' -Fail -Detail "got $($rows.Count) rows"; $script:failed++ }
} catch { Write-TestCase 'accepts-empty-session' -Fail -Detail $_.Exception.Message; $script:failed++ }

# A MISSING file means "never armed" -- an empty result and no throw.
try {
    $rows = @(Read-TraceFile (Join-Path $dir 'never-existed.csv'))
    if ($rows.Count -eq 0) { Write-TestCase 'missing-file-is-empty-not-fault' -Pass }
    else { Write-TestCase 'missing-file-is-empty-not-fault' -Fail; $script:failed++ }
} catch { Write-TestCase 'missing-file-is-empty-not-fault' -Fail -Detail $_.Exception.Message; $script:failed++ }

# ---- the refusal paths -----------------------------------------------------
Expect-Refusal 'refuses-renamed-column' 'expected:' `
    (New-TraceFile 'renamed.csv' (@($hdr -replace 'typetext', 'sig') + $good))

Expect-Refusal 'refuses-reordered-columns' 'expected:' `
    (New-TraceFile 'reordered.csv' (@($hdr -replace '^seq,input', 'input,seq') + $good))

Expect-Refusal 'refuses-extra-column' 'expected:' `
    (New-TraceFile 'extra.csv' (@("$hdr,extra") + $good))

Expect-Refusal 'refuses-unknown-kind' 'unknown kind' `
    (New-TraceFile 'kind.csv' (@($hdr) + $good + @('6,6,entry2,XLL,9,,,17248,2759222880000,X.xll,F,F,Q,,,,,,,,,')))

Expect-Refusal 'refuses-unknown-source' 'unknown source' `
    (New-TraceFile 'source.csv' (@($hdr) + $good + @('6,6,entry,COM,9,,,17248,2759222880000,X.xll,F,F,Q,,,,,,,,,')))

Expect-Refusal 'refuses-seq-inversion' 'strictly increasing' `
    (New-TraceFile 'inversion.csv' (@($hdr) + $good[0], $good[2], $good[1]))

Expect-Refusal 'refuses-duplicate-seq' 'strictly increasing' `
    (New-TraceFile 'dupe.csv' (@($hdr) + $good[0], $good[0]))

Expect-Refusal 'refuses-nonnumeric-qpc' 'non-numeric qpc' `
    (New-TraceFile 'qpc.csv' (@($hdr) + @('1,1,entry,XLL,9,,,17248,soon,X.xll,F,F,Q,,,,,,,,,')))

Expect-Refusal 'refuses-torn-row' 'columns|non-numeric|unknown kind|bad input' `
    (New-TraceFile 'torn.csv' (@($hdr) + $good[0], '2,ex'))

# A ROW SHORT BY EXACTLY ONE COLUMN is the one that gets through everything
# else. Import-Csv pads it with $nulls that read as empty fields, so every
# value after the missing one is silently off by one and the row looks merely
# blank -- which is a confident wrong answer, not a gap.
Expect-Refusal 'refuses-row-short-by-one' 'columns' `
    (New-TraceFile 'short.csv' (@($hdr) + ($good[1] -replace ',604,exit$', ',604')))

Expect-Refusal 'refuses-row-long-by-one' 'columns' `
    (New-TraceFile 'long.csv' (@($hdr) + ($good[1] + ',spare')))

Expect-Refusal 'refuses-duplicate-input' 'input .* used twice' `
    (New-TraceFile 'dupein.csv' (@($hdr) + $good[0] + @('2,1,exit,XLL,9,,,17248,2759222880000,X.xll,F,F,Q,,,,,,,,,')))

Expect-Refusal 'refuses-unknown-format' 'no reader' `
    (New-TraceFile 'trace.json' @('{"traceEvents":[]}'))

if ($script:failed -eq 0) { Complete-Test -Pass }
else { Complete-Test -Fail -Detail "$script:failed contract case(s) failed" }
