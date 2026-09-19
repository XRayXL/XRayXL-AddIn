# A macro Excel runs while VBA sits in DoEvents is a new top-level chain.
#
# An OnTime macro fires inside SlowWork's DoEvents loop. The interpreter stack says it is on top
# of SlowWork, but SlowWork did not call it: it must read depth 1 with no parent, and its own
# callee must nest under it.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public gTicked As Boolean

Public Sub RunReentry()
    gTicked = False
    Application.OnTime Now + TimeSerial(0, 0, 1), "TimerMac"
    SlowWork
End Sub

Private Sub SlowWork()
    Dim t As Single
    t = Timer
    Do While Timer - t < 3 And Not gTicked
        DoEvents
    Loop
End Sub

Public Sub TimerMac()
    gTicked = True
    Helper
End Sub

Private Sub Helper()
    Dim n As Long
    n = 1
End Sub

Public Function Ticked() As Boolean
    Ticked = gTicked
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx

    New-XRayMacroBook $sx 'DoEventsChain' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Leaf $leaf -Body { $app.Run($leaf + '!RunReentry') | Out-Null; [bool]$app.Run($leaf + '!Ticked') }
    if (-not $s.Result) { Complete-Test -Fail -Detail 'the timer never fired inside DoEvents; nothing was tested' }
    $t = ConvertFrom-XRayTotals $s.Totals

    $slow   = EntryRowOf $s.Rows 'SlowWork'
    $timer  = EntryRowOf $s.Rows 'TimerMac'
    $helper = EntryRowOf $s.Rows 'Helper'
    Check 'the-timer-macro-is-top-level' (($timer.depth -eq '1') -and ($timer.parent -eq '0')) "TimerMac depth $($timer.depth) parent $($timer.parent)"
    Check 'its-callee-nests-under-it' (($helper.depth -eq '2') -and ($helper.parent -eq $timer.span)) "Helper depth $($helper.depth) parent $($helper.parent), TimerMac span $($timer.span)"
    Check 'the-interrupted-chain-keeps-its-depth' ($slow.depth -eq '2') "SlowWork depth $($slow.depth)"
    Check 'one-chain-counted' ($t.doEventsChains -eq 1) "$($s.Totals)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "TimerMac depth 1 parent 0; doEventsChains=$($t.doEventsChains)"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
