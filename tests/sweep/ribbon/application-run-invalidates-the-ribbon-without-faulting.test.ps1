# Drives arming and the settings from Application.Run while the ribbon add-in is connected. Each change
# makes the ribbon call IRibbonUI::Invalidate, a call back into Office; the assertion is that no COM
# entry point faulted. A contained fault is a line in the LOG at ERROR, not a crash file.
#
# Sessions here are hidden, and Excel builds the ribbon only where there is a window, so this normally
# exercises the notification path and not the call into Office. The verdict says which run it was.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_ribbon_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $verdict = Get-RibbonLogVerdict $paths.Log
    if ($verdict -ne 'loaded') {
        Complete-Test -Skip -Detail "no live ribbon to drive: $verdict"
    }
    Check 'connected-before' (Get-RibbonConnected $app) 'the add-in was not connected at the start'

    # which run is this? recorded before anything is driven
    $uiLive = Test-RibbonUiLive $paths.Log
    Write-XRayObservation 'ribbon-ui-built' $(if ($uiLive) {
        'onLoad fired: IRibbonUI held, so Invalidate really calls into Office'
    } else {
        'no onLoad in this session (no window), so Invalidate is a no-op and only the notification path is exercised'
    })

    # ---- settings: each accepted set fires the notification ----
    foreach ($pass in 1..2) {
        foreach ($src in @('XLL', 'VBA')) {
            foreach ($depth in @('OFF', 'TOP', 'ALL')) {
                $echo = Set-XRayTraceParam $sx $src 'DEPTH' $depth
                Check "set-$src-depth-$depth-pass$pass" ($echo -notmatch 'refused') $echo
            }
            foreach ($name in @('ARGS', 'RETVAL', 'OBJECTS')) {
                [void](Set-XRayTraceParam $sx $src $name 'FALSE')
                [void](Set-XRayTraceParam $sx $src $name 'TRUE')
            }
        }
    }
    Check 'connected-after-settings' (Get-RibbonConnected $app) `
          'the add-in went away while settings were being changed'

    # ---- arming, changing under a ribbon that did not ask ----
    foreach ($src in @('XLL', 'VBA')) { [void](Set-XRayTraceParam $sx $src 'DEPTH' 'ALL') }
    foreach ($round in 1..6) {
        $mark = Get-LogLength $paths.Log
        [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
        $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
        Check "armed-round-$round" ([bool]$armLine) 'the arm never reported'
        Check "isarmed-round-$round" ([bool]$app.Run('XRayXL_IsArmed')) 'IsArmed disagreed with the arm'
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        Check "disarmed-round-$round" (-not [bool]$app.Run('XRayXL_IsArmed')) 'still armed after the disarm'
    }

    Check 'connected-after-arming' (Get-RibbonConnected $app) `
          'the add-in went away while arming and disarming'

    # ---- did any COM entry point fault? read from the log, which always exists ----
    $faults = @(Get-Content $paths.Log -ErrorAction SilentlyContinue |
                Select-String -Pattern 'RIBBON COM METHOD FAULTED' |
                ForEach-Object { $_.ToString().Trim() })
    Check 'no-com-entry-point-faulted' ($faults.Count -eq 0) ($faults -join ' | ')

    # the settings must still be right afterwards
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'TOP')
    $back = Get-XRayTraceParam $sx 'XLL' 'DEPTH'
    Check 'settings-survive-the-invalidates' ($back -match 'TOP') "GetTraceParam said '$back'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    $what = if ($uiLive) { 'invalidated a built ribbon' } else { 'drove the notification path (no ribbon built: no window)' }
    Complete-Test -Pass -Detail "state driven from Application.Run $what with no fault and no loss"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
