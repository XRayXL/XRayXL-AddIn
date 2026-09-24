# Registering XRayXL again in an Excel that has it loaded keeps its session: Excel calls xlAutoOpen
# again, and the log, crash handlers and pin are per-process. Reopening the log would truncate the
# session's record, and installing the crash filter again would make it call itself.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId
    $app = $sx.App

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark
    [void](Invoke-XRayDisarm $sx)
    $before = [IO.File]::ReadAllText($paths.Log)

    $xll = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\build\addin\XRayXL64.xll')).Path
    $registered = [bool]$app.RegisterXLL($xll)
    $after = [IO.File]::ReadAllText($paths.Log)
    $opened = ([regex]::Matches($after, 'XRayXL log opened')).Count

    $mark2 = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armAgain = Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark2
    [void](Invoke-XRayDisarm $sx)

    Check 'the-first-arm-worked' ($armLine -match 'armed \d+ of') "$armLine"
    Check 'excel-registered-it-again' $registered "RegisterXLL returned $registered"
    Check 'the-log-still-holds-the-earlier-session' ($after.StartsWith($before)) `
          "log was $($before.Length) chars before registering again, $($after.Length) after"
    Check 'the-log-was-opened-once' ($opened -eq 1) "'XRayXL log opened' lines: $opened"
    Check 'it-still-arms-afterwards' ($armAgain -match 'armed \d+ of') "$armAgain"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'registering again kept the log and still arms'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
