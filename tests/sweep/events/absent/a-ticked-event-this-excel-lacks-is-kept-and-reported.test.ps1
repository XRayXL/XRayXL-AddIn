# A ticked event this Excel lacks is kept, so settings carry between machines, and the arm log names it.
# The session's diagnostic switch makes Excel appear to lack SheetTableUpdate and WorkbookModelChange.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $log = (Get-XRayPaths $sx.ProcId).Log

    $mark = Get-LogLength $log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $line = Wait-LogLine $log 'events: ' $mark
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { throw $lossy }

    Check 'the-arm-log-names-the-missing-events' ($line -match 'not in this Excel: SheetTableUpdate,WorkbookModelChange') "log '$line'"
    Check 'the-rest-still-reads-as-calc' ($line -match 'events: recording Calc;') "log '$line'"
    $kept = [string]$sx.App.Run('XRayXL_GetTraceParam', 'EVENTS', 'SheetTableUpdate')
    Check 'the-setting-is-kept' ($kept -eq 'TRUE') "reads '$kept'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail $line
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
