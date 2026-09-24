# Disarm while an add-in thread writes rows through the ring: the ring is freed at disarm, so it must
# outlive every producer that saw the session armed. The smallest ring makes the worker wait for
# room, which widens the window.
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

    $app.Run('TxHammerStart') | Out-Null
    $hammerOn = $true
    $armed = 0
    $cycles = 0
    foreach ($size in '64', '16KB') {
        [void](Set-XRayTraceParam $sx $null 'BUFFERSIZE' $size)
        for ($i = 0; $i -lt 20; $i++) {
            $cycles++
            $mark = Get-LogLength $paths.Log
            [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
            $line = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
            if ($line -match 'armed \d+ of') { $armed++ }
            else { Stop-Hammer; Complete-Test -Fail -Detail "cycle $cycles ($size ring) did not arm: $line" }
            $before = [double]$app.Run('TxHammerCalls')
            [void](Wait-XRayCondition { [double]$app.Run('TxHammerCalls') -gt $before + 200 } 10)
            [void](Invoke-XRayDisarm $sx)
        }
    }
    Stop-Hammer
    $calls = [double]$app.Run('TxHammerCalls')

    Check 'every-cycle-armed' ($armed -eq $cycles) "$armed of $cycles"
    Check 'the-worker-kept-calling' ($calls -gt 0) "calls=$calls"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "$cycles arm/disarm cycles through the ring under a calling worker; $calls calls"
}
catch {
    Stop-Hammer
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
