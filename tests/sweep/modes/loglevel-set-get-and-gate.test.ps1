# LOGLEVEL is source-less like BUFFERSIZE but controls the log, not the trace: it must be settable
# while armed, and must visibly gate the log, since an unobservable setter reads like a working one.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # A book with an XLL UDF, so XRayXL_Arm hooks something and the session is genuinely armed.
    New-XRayMacroBook $sx 'LogLevel' -Cells @{ 'A1' = '=TxB(2,3)' } -Format xlsx

    # ---- Set / Get echo --------------------------------------------------
    $e = [string]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'DEBUG')
    Check 'set-echoes-the-level' ($e -match 'LOGLEVEL=DEBUG') $e
    $g = [string]$app.Run('XRayXL_GetTraceParam', 'LOGLEVEL')
    Check 'get-reads-it-back' ($g -eq 'DEBUG') "GetTraceParam(LOGLEVEL)='$g'"

    # ---- an unknown level is refused, and changes nothing ----------------
    $bad = [string]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'LOUD')
    Check 'invalid-level-refused' ($bad -match '#Err') $bad
    $g2 = [string]$app.Run('XRayXL_GetTraceParam', 'LOGLEVEL')
    Check 'refusal-changed-nothing' ($g2 -eq 'DEBUG') "still '$g2' after a refused set"

    # ---- settable while armed (the point vs DEPTH/ARGS) ------------------
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    [void](Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark)
    $ea = [string]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'WARNING')
    Check 'settable-while-armed' (($ea -match 'LOGLEVEL=WARNING') -and ($ea -notmatch '#Err')) $ea
    $ga = [string]$app.Run('XRayXL_GetTraceParam', 'LOGLEVEL')
    Check 'reads-while-armed' ($ga -eq 'WARNING') "while armed='$ga'"
    [void](Invoke-XRayDisarm $sx)

    # ---- gating: at ERROR an INFO action does not reach the log; at INFO it
    # does. SetTraceParam logs "function: ..." at INFO, so it is the probe.
    # The log is written synchronously; the one-second window only catches a late line.
    [void]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'ERROR')
    $m1 = Get-LogLength $paths.Log
    [void]$app.Run('XRayXL_SetTraceParam', 'XLL', 'DEPTH', 'ALL')   # an INFO action
    $grewAtError = Wait-XRayCondition { (Get-LogLength $paths.Log) -gt $m1 } 1
    Check 'info-action-is-gated-at-error' (-not $grewAtError) "log grew=$grewAtError at ERROR after an INFO action"

    [void]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'INFO')
    $m2 = Get-LogLength $paths.Log
    [void]$app.Run('XRayXL_SetTraceParam', 'XLL', 'DEPTH', 'ALL')
    $grewAtInfo = Wait-XRayCondition { (Get-LogLength $paths.Log) -gt $m2 } 10
    Check 'info-action-reaches-the-log-at-info' $grewAtInfo "log grew=$grewAtInfo at INFO after an INFO action"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "LOGLEVEL: Set/Get, invalid refused, settable while armed, gates INFO at ERROR"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
