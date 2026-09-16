# A PROCEDURE NAME LONGER THAN THE REGISTRATION CAPTURE.
#
# The watch keeps 127 characters of a procedure name. Cut, a long name can spell
# a shorter export exactly, and hooking that one would trace an unrelated
# function under the long one's name. TracedAddin registers such a name after
# arming, then TxCallsPrefix calls the shorter export the cut name spells.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    # Registered by an earlier run in this Excel, it is hooked whole at arm and the watch never sees it cut.
    if (@($app.Evaluate('TxLongNamed(1)'))[0] -is [double]) {
        Complete-Test -Skip -Detail 'TxLongNamed was registered before this arm, so the watch cannot decline it'
    }
    # The declined count is cumulative for the process, so this run's decline is the rise from here.
    $cutPattern = '(\d+) declined \(a field cut short\)'
    $declinedBefore = 0
    $last = @(Select-String -Path $paths.Log -Pattern $cutPattern -ErrorAction SilentlyContinue) | Select-Object -Last 1
    if ($last) { $declinedBefore = [int]$last.Matches[0].Groups[1].Value }

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    if ($armLine -notmatch 'armed \d+ of') { Complete-Test -Fail -Detail "arm: $armLine" }
    $watchLine = Wait-LogLine $paths.Log 'register watch: (installed|NOT installed|DISABLED)' $mark 15
    if ($watchLine -notmatch 'installed on MdCallBack12') { Complete-Test -Fail -Detail "watch: $watchLine" }

    $mark2 = Get-LogLength $paths.Log
    $app.Run('TxRegisterLongName') | Out-Null
    # A declined capture logs nothing, so this waits for a hook that should not happen, briefly.
    $applied = Wait-LogLine $paths.Log 'register watch: applied' $mark2 5
    Write-XRayObservation 'watch-after-long-name' "$applied"
    $registered = [string]$app.Evaluate('TxLongNamed(1)')

    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxCallsPrefix(2)'
    Invoke-XRayRecalc $app
    $value = Get-XRayCellText $ws.Cells.Item(1, 1)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $declinedAfter = $declinedBefore
    $last = @(Select-String -Path $paths.Log -Pattern $cutPattern) | Where-Object { $_.LineNumber -gt $mark } | Select-Object -Last 1
    if ($last) { $declinedAfter = [int]$last.Matches[0].Groups[1].Value }

    $rows = @(Read-TraceFile (Get-XRayTraceCsv $sx.ProcId))
    # Told apart by what returned: the long export gives x+0.75, the short one x+0.5.
    $misnamed = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.function -eq 'TxLongNamed' -and $_.ret -ne '1.75' })
    $caller = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'TxCallsPrefix' })

    Check 'excel-registered-the-long-name' ($registered -eq '1.75') "TxLongNamed(1)='$registered'"
    Check 'the-caller-answered' ($value -eq '2.5') "A1='$value'"
    Check 'the-caller-was-traced' ($caller.Count -ge 1) "TxCallsPrefix entries: $($caller.Count)"
    Check 'the-watch-declined-the-cut-name' ($declinedAfter -gt $declinedBefore) `
          "declined (a field cut short): $declinedBefore before this arm, $declinedAfter after"
    Check 'no-row-carries-the-cut-name' ($misnamed.Count -eq 0) `
          "TxLongNamed exits not returning the long export's 1.75: $($misnamed.Count)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'a cut procedure name hooked nothing'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
