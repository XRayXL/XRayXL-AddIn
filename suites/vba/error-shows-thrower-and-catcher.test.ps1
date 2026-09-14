# WHERE WAS IT THROWN, AND WHO CAUGHT IT.
#
# Without an outcome, an error unwind and a clean return give identical exit
# rows, and a fully unhandled unwind (no exit opcodes) reads as a slow,
# successful call.
#
# The exit row carries its own `outcome` column, and a chain reads outwards from
# the throwing frame:
#
#     Thrower   outcome threw       the raise happened HERE
#     Middle    outcome unwound     it ran nothing after the raise -- passed through
#     Outer     outcome handled     it ran again, so it caught it
#
# HOW THE RULE WORKS, and why it is not the obvious one. Slot 497 is the raise
# opcode. It fires FOUR TIMES per Err.Raise, so the recorder dedupes. The error clears on the SECOND statement to reach a frame,
# never the first: both a handler and an unwind run a statement in the raising
# frame immediately after the raise, through literally the same opcode, so
# clearing on the first reported the thrower as `returned` and the frame the
# error merely passed through as errored -- exactly backwards.
#
# ORDER MATTERS IN THE ASSERTIONS. `unwound` is the one that cannot be faked by
# a coincidence: it means a frame closed while an error was in flight AND had
# run nothing since, which no clean return can produce.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
' Outer catches; Middle only passes it through; Thrower raises.
Public Sub E_Outer()
    Dim n As Long
    On Error GoTo Caught
    n = 1
    E_Middle
    n = 2
    Exit Sub
Caught:
    n = 3            ' <- the statement that proves Outer resumed
    n = 4
End Sub

Public Sub E_Middle()
    Dim m As Long
    m = 1
    E_Thrower
    m = 2            ' never reached
End Sub

Public Sub E_Thrower()
    Dim t As Long
    t = 1
    Err.Raise 5, "XRayCase", "a deliberate error"
    t = 2            ' never reached
End Sub

' A clean chain, so `returned` is asserted against something in the same run.
Public Sub E_Clean()
    Dim q As Long
    q = 1
    E_CleanInner
    q = 2
End Sub
Public Sub E_CleanInner()
    Dim r As Long
    r = 1
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ErrChain' @(
        @{ Kind=1; Name='ErrCase'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    # THE RAISE SLOT HAS TO BE PATCHED, or every outcome reads `returned` and
    # the assertions below would fail for a reason that is not this feature.
    $raiseOk = ($armLine -match '(\d+) raise') -and ([int]$Matches[1] -ge 1)
    Check 'raise-slot-patched' $raiseOk "arm line: $armLine"
    # an unverified raise slot is a product failure, not a SKIP: every outcome would read `returned`
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail ("the raise slot did not verify on this VBE7, so no outcome " +
                                     "below can be attributed: $armLine")
    }

    $app.Run($leaf + '!E_Outer') | Out-Null
    $app.Run($leaf + '!E_Clean') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $exits = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })

    $t = Outcome $rows 'E_Thrower'; $m = Outcome $rows 'E_Middle'; $o = Outcome $rows 'E_Outer'
    $c = Outcome $rows 'E_Clean';   $ci = Outcome $rows 'E_CleanInner'

    $absent = @(@('E_Thrower', $t), @('E_Middle', $m), @('E_Outer', $o) |
                Where-Object { $_[1] -eq '(no row)' } | ForEach-Object { $_[0] })
    Check 'all-three-frames-traced' ($absent.Count -eq 0) `
          "no exit row for: $($absent -join ',') -- thrower='$t' middle='$m' outer='$o'"
    Check 'thrower-says-threw'   ($t -eq 'threw')   "E_Thrower outcome='$t'"
    Check 'passthrough-unwound'  ($m -eq 'unwound') "E_Middle outcome='$m'"
    Check 'catcher-says-handled' ($o -eq 'handled') "E_Outer outcome='$o'"

    # THE NEGATIVE CONTROL. Without it, a bug that stamped every row `threw`
    # would satisfy the first assertion and look like a pass.
    Check 'clean-calls-say-returned' (($c -eq 'returned') -and ($ci -eq 'returned')) `
          "E_Clean='$c' E_CleanInner='$ci'"

    # Every exit row must carry an outcome.
    $missing = @($exits | Where-Object { -not $_.outcome })
    Check 'every-exit-row-carries-an-outcome' ($missing.Count -eq 0) `
          "rows without outcome: $($missing.Count) of $($exits.Count)"


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "chain: E_Thrower=$t -> E_Middle=$m -> E_Outer=$o; clean=$c/$ci"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
