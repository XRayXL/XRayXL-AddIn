# The register watch's disarm line counts THIS arming session, like every other counter.
#
# Its counts were process-wide, so a session that registered nothing reported the registrations
# an earlier session had seen: "6 were xlfRegister, 6 hooked" in an arm with no registration.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

function Watch-Line($paths, $mark) {
    Wait-LogLine $paths.Log 'register watch: \d+ C API call\(s\) seen' $mark 20
}

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App
    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')
    $ws = $app.ActiveSheet
    $ws.Cells.Item(1, 1).Formula = '=TxB(2,3)'

    # Session one registers a function while armed.
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $app.Run('TxRegisterNoResult') | Out-Null
    Invoke-XRayRecalc $app
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $first = Watch-Line $paths $mark

    # Session two registers nothing.
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    Invoke-XRayRecalc $app
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $second = Watch-Line $paths $mark

    $n1 = if ($first -match '(\d+) were xlfRegister') { [int]$Matches[1] } else { -1 }
    $n2 = if ($second -match '(\d+) were xlfRegister') { [int]$Matches[1] } else { -1 }
    Check 'session-one-saw-its-registration' ($n1 -ge 1) "$first"
    Check 'session-two-reports-none' ($n2 -eq 0) "$second"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "session one: $n1 registration(s); session two: $n2"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
