# The outcome column says where an error was thrown and who caught it; without it an unwind and
# a clean return give identical exit rows.
#
#      Thrower   outcome threw       the raise happened here
#      Middle    outcome unwound     it ran nothing after the raise
#      Outer     outcome handled     it ran again, so it caught it
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
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

    # the negative control: a bug stamping every row `threw` would otherwise pass
    Check 'clean-calls-say-returned' (($c -eq 'returned') -and ($ci -eq 'returned')) `
          "E_Clean='$c' E_CleanInner='$ci'"

    $missing = @($exits | Where-Object { -not $_.outcome })
    Check 'every-exit-row-carries-an-outcome' ($missing.Count -eq 0) `
          "rows without outcome: $($missing.Count) of $($exits.Count)"


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "chain: E_Thrower=$t -> E_Middle=$m -> E_Outer=$o; clean=$c/$ci"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
