# The ribbon add-in connects, and its registration does not outlive the connect: the CLSID, ProgId and
# Addins entry are written, used and deleted at once, and the entry is LoadBehavior=0 so nothing autoloads.
# Both are asserted together: no keys and no add-in is a ribbon that never loaded.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_ribbon_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $verdict = Get-RibbonLogVerdict $paths.Log
    $connected = Get-RibbonConnected $app

    # a fail-soft is reported with its reason; the add-in working without the ribbon is the other suites' job
    if ($verdict -like 'failsoft*') {
        Complete-Test -Fail -Detail "the ribbon did not load: $verdict"
    }
    elseif ($verdict -eq 'off') {
        Complete-Test -Skip -Detail 'XRAYXL_RIBBON=0 in this session'
    }

    Check 'log-says-the-ribbon-loaded' ($verdict -eq 'loaded') "log verdict was '$verdict'"
    Check 'excel-agrees-it-is-connected' $connected `
          "COMAddIns('$script:RibbonProgId').Connect was '$connected'"

    # sampled: the suite runs in parallel and the registry is per-user, so a key seen once may be another worker's
    $left = Get-RibbonKeysLeftBehind
    Check 'no-registration-outlives-the-connect' ($left.Count -eq 0) `
          ("still present: " + ($left -join '; '))

    # an Addins entry at LoadBehavior=3 would autoload the add-in into the next Excel
    $addins = 'HKCU:\Software\Microsoft\Office\Excel\Addins\XRayXL.RibbonUI'
    $lb = $null
    if (Test-Path $addins) { $lb = (Get-ItemProperty $addins -ErrorAction SilentlyContinue).LoadBehavior }
    Check 'no-autoloading-addins-entry' ($null -eq $lb -or $lb -eq 0) `
          "LoadBehavior was '$lb' (3 would autoload us into the next Excel)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'the ribbon is connected and nothing of its registration remains'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
