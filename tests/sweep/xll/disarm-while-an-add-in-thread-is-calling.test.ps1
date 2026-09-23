# DISARM WHILE A THREAD OF THE ADD-IN'S OWN IS CALLING ONE OF ITS EXPORTS.
#
# Excel is not the only caller of an XLL. TracedAddin runs a worker that calls
# TxHammered in a loop, and the session is armed and disarmed under it many
# times. A detour torn down while that thread is inside it leaves the thread to
# call a trampoline that no longer exists, so Excel surviving is the assertion.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# A running worker keeps Excel from closing, so every way out stops it first.
$hammerOn = $false
function Stop-Hammer {
    if ($script:hammerOn) { try { [void]$app.Run('TxHammerStop') } catch {}; $script:hammerOn = $false }
}

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')
    # The worker is not an Excel thread, so it must not ask Excel who called.
    # Synchronous output, so this measures the detours and not the output ring.
    [void](Set-XRayTraceParam $sx $null 'BUFFERSIZE' '0')

    $app.Run('TxHammerStart') | Out-Null
    $hammerOn = $true
    $cycles = 40
    $armed = 0
    for ($i = 0; $i -lt $cycles; $i++) {
        $mark = Get-LogLength $paths.Log
        [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
        $line = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
        if ($line -match 'armed \d+ of') { $armed++ }
        else { Stop-Hammer; Complete-Test -Fail -Detail "cycle $($i + 1) did not arm: $line" }
        # Let the worker run through the detour before tearing it down.
        $before = [double]$app.Run('TxHammerCalls')
        [void](Wait-XRayCondition { [double]$app.Run('TxHammerCalls') -gt $before + 200 } 10)
        [void](Invoke-XRayDisarm $sx)
    }
    Stop-Hammer
    $calls = [double]$app.Run('TxHammerCalls')

    # Raw lines: a call in flight at disarm legitimately leaves an entry with no exit.
    $csv = Get-XRayTraceCsv $sx.ProcId
    $hammerRows = @([IO.File]::ReadAllLines($csv) | Where-Object { $_ -match ',TxHammered,' }).Count

    Check 'every-cycle-armed' ($armed -eq $cycles) "$armed of $cycles"
    Check 'the-worker-kept-calling' ($calls -gt 0) "calls=$calls"
    Check 'the-last-session-traced-the-worker' ($hammerRows -gt 0) "rows naming TxHammered in $csv : $hammerRows"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "$cycles arm/disarm cycles under a calling worker; $calls calls"
}
catch {
    Stop-Hammer
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
