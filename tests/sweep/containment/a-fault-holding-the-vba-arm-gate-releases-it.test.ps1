# A fault contained part-way through a VBA arm or disarm must release the arm gate: a contained fault
# runs no destructor, and a gate left held refuses every later arm and disarm. XRayXL_FaultProbe
# "ARMGATE" faults while holding it (needs XRAYXL_DIAG).
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    # VBA must be loaded for the VBA side to arm at all.
    New-XRayMacroBook $sx 'ArmGate' @(@{ Kind = 1; Name = 'M'; Code = "Public Function G_One() As Long`r`n    G_One = 1`r`nEnd Function" })

    $mark = Get-LogLength $paths.Log
    $probe = $null
    try { $probe = $app.Run('XRayXL_FaultProbe', 'ARMGATE') } catch { $probe = "threw: $($_.Exception.Message)" }
    Write-Output "probe returned: $probe"
    Check 'the-fault-was-contained' (($probe -is [int]) -or ($probe -is [double])) "XRayXL_FaultProbe returned '$probe'"
    $released = Wait-LogLine $paths.Log 'its gate is released' $mark 10
    Check 'the-release-is-logged' ([bool]$released) 'no "its gate is released" line after the probe'

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ARMED|arm/disarm already in progress|VBA tracing: refused' $mark 30
    Check 'vba-arms-after-the-fault' (($pressed -eq 'pressed') -and ($armLine -match 'VBA tracing: ARMED')) "$pressed / $armLine"

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Disarm')
    $disarmLine = Wait-LogLine $paths.Log 'VBA tracing: disarmed|arm/disarm already in progress' $mark 30
    Check 'vba-disarms-after-the-fault' ($disarmLine -match 'VBA tracing: disarmed') "$disarmLine"
    Check 'nothing-left-armed' (-not [bool]$app.Run('XRayXL_IsArmed')) 'XRayXL_IsArmed still TRUE after disarm'

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'a fault holding the arm gate left VBA free to arm and disarm'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
