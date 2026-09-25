# Disarm from Excel while a macro waits in DoEvents: the macro is still running, so it reads
# `returned`, never `abandoned` like a frame the error dialog's End left behind.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub WaitForDisarm()
    Dim t As Single
    Application.OnTime Now + TimeSerial(0, 0, 1), "XRayXL_Disarm"
    t = Timer
    Do While Timer - t < 4
        DoEvents
    Loop
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'DisarmInWait' @(@{ Kind=1; Name='M'; Code=$moduleCode })
    $leaf = (Get-XRayMacroBook).Leaf
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    [void](Wait-LogLine $paths.Log 'armed \d+ of|nothing armed|could not|XLL tracing: OFF' $mark)

    # Excel runs the disarm from its message loop, with no VBA between it and the waiting macro.
    $app.Run($leaf + '!WaitForDisarm') | Out-Null
    [void](Wait-XRayTraceClosed $sx.ProcId)

    $all = @(Read-TraceRows $sx.ProcId)
    $rows = @(Select-BookRows $all $leaf)
    Check 'excel-ran-the-disarm' ($all.Count -and $all[-1].function -eq 'disarm') `
          "last row: $(if ($all.Count) { $all[-1].function })"
    $x = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq 'WaitForDisarm' })
    Check 'the-waiting-macro-reads-running' `
          (($x.Count -eq 1) -and ($x[0].outcome -eq 'returned') -and ($x[0].trust -eq 'flush')) `
          ("WaitForDisarm=" + (($x | ForEach-Object { "$($_.outcome)/$($_.trust)" }) -join ','))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'a macro in DoEvents at disarm reads returned/flush'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
