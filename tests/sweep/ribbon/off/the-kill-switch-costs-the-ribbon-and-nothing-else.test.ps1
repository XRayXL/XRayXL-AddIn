# XRAYXL_RIBBON=0, the kill switch: a session with no COM object of ours at all, which is what makes
# "was it the ribbon?" answerable. It must cost nothing else: no buttons, no CLSID, ProgId or Addins
# entry, and every command and setter working as it does with them.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '..\_ribbon_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $verdict = Get-RibbonLogVerdict $paths.Log
    Check 'the-switch-was-seen' ($verdict -eq 'off') `
          "expected the log to record XRAYXL_RIBBON=0; it said '$verdict'"
    Check 'no-add-in-connected' (-not (Get-RibbonConnected $app)) `
          'COMAddIns still reports the ribbon add-in connected'

    # never written, rather than written and removed
    $left = Get-RibbonKeysLeftBehind
    Check 'nothing-registered' ($left.Count -eq 0) ("present: " + ($left -join '; '))

    # ---- and the add-in is otherwise entirely itself -----------------------
    $echo = Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'TOP'
    Check 'setter-works' ($echo -notmatch 'refused') $echo
    $back = Get-XRayTraceParam $sx 'XLL' 'DEPTH'
    Check 'getter-agrees' ([string]$back -match 'TOP') "GetTraceParam said '$back'"

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'ALL')
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    Check 'arm-works' ([bool]$armLine) 'the arm never reported'
    Check 'isarmed-works' ([bool]$app.Run('XRayXL_IsArmed')) 'IsArmed said false after arming'

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    Check 'disarm-works' (-not [bool]$app.Run('XRayXL_IsArmed')) 'still armed after the disarm'

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'no ribbon buttons, no COM object, no registration -- and every command unaffected'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
