# LOGLEVEL is a Get/Set parameter like the others, source-less like BUFFERSIZE (the
# word lands in the Source slot, the level after it), but with two things it must
# PROVE:
#   * it is SETTABLE WHILE ARMED -- it controls the log, not the trace, so the
#     armed-refusal that guards DEPTH/ARGS must NOT apply to it;
#   * setting it actually GATES the log. A setter whose effect cannot be
#     observed reads exactly like one that works, so this reads the LOG and
#     checks an INFO action vanishes at level ERROR and returns at INFO.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # A book with an XLL UDF, so XRayXL_Arm actually hooks something and the
    # session is genuinely ARMED for the settable-while-armed check.
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

    # ---- SETTABLE WHILE ARMED (the point vs DEPTH/ARGS) ------------------
    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    [void](Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not' $mark)
    $ea = [string]$app.Run('XRayXL_SetTraceParam', 'LOGLEVEL', 'WARNING')
    Check 'settable-while-armed' (($ea -match 'LOGLEVEL=WARNING') -and ($ea -notmatch '#Err')) $ea
    $ga = [string]$app.Run('XRayXL_GetTraceParam', 'LOGLEVEL')
    Check 'reads-while-armed' ($ga -eq 'WARNING') "while armed='$ga'"
    [void](Invoke-XRayDisarm $sx)

    # ---- GATING: at ERROR an INFO action does not reach the log; at INFO it
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
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
