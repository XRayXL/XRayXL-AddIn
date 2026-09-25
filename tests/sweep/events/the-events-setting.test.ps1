# EVENTS is set and read one event at a time, refuses a name Excel does not have, and is fixed while armed.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

function Get-Event($Sx, $Name) {
    if ($null -eq $Name) { return [string]$Sx.App.Run('XRayXL_GetTraceParam', 'EVENTS') }
    return [string]$Sx.App.Run('XRayXL_GetTraceParam', 'EVENTS', $Name)
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $log = (Get-XRayPaths $sx.ProcId).Log

    Check 'the-default-is-calc' ((Get-Event $sx $null) -eq 'Calc') "reads '$(Get-Event $sx $null)'"
    Check 'a-calc-event-reads-true' ((Get-Event $sx 'SheetChange') -eq 'TRUE') "reads '$(Get-Event $sx 'SheetChange')'"
    Check 'another-reads-false' ((Get-Event $sx 'SheetSelectionChange') -eq 'FALSE') "reads '$(Get-Event $sx 'SheetSelectionChange')'"

    $typo = Set-XRayTraceParam $sx 'EVENTS' 'SheetChnage' $true
    Check 'an-unknown-name-is-refused' ($typo -like '#Err*') "echo '$typo'"
    Check 'a-refusal-changes-nothing' ((Get-Event $sx $null) -eq 'Calc') "reads '$(Get-Event $sx $null)'"

    # Second, as the other source-less words may go: read back, but refused plainly by the setter.
    $second = [string]$app.Run('XRayXL_GetTraceParam', [Type]::Missing, 'EVENTS')
    Check 'second-reads-the-preset' ($second -eq 'Calc') "reads '$second'"
    $wrongSlot = [string]$app.Run('XRayXL_SetTraceParam', [Type]::Missing, 'EVENTS', 'SheetSelectionChange')
    Check 'second-is-refused-by-saying-where-it-goes' ($wrongSlot -like '#Err - EVENTS comes first*') "echo '$wrongSlot'"

    $echo = Set-XRayTraceParam $sx 'EVENTS' 'sheetselectionchange' $true
    Check 'names-are-case-insensitive' ($echo -like 'EVENTS SheetSelectionChange=TRUE*') "echo '$echo'"
    Check 'a-mix-reads-custom' ((Get-Event $sx $null) -eq 'Custom') "reads '$(Get-Event $sx $null)'"

    $mark = Get-LogLength $log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $line = Wait-LogLine $log 'events: ' $mark
    Check 'the-arm-log-names-what-is-recorded' ($line -match 'events: recording .*SheetSelectionChange') "log '$line'"
    $armed = Set-XRayTraceParam $sx 'EVENTS' 'WindowResize' $true
    Check 'refused-while-armed' ($armed -like '#Err - cannot change settings while armed*') "echo '$armed'"
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { throw $lossy }
    Check 'the-armed-refusal-changed-nothing' ((Get-Event $sx 'WindowResize') -eq 'FALSE') "reads '$(Get-Event $sx 'WindowResize')'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail $line
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
