# A real error swallowed by On Error Resume Next in the same frame.
#
# A benign object-model raise and a real error caught in place are indistinguishable at the
# raise (same opcode, 497, same registers), and both let the frame run on to its own exit, so
# both read `returned`. This pins that, and fails loudly on `threw`, which would mean the tracer
# thought the error escaped.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub P2_ResumeNext()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    On Error Resume Next
    Dim z As Long
    z = 1 \ 0                        ' a REAL error (division by zero), swallowed here
    ws.Range("A1").Value = "after"   ' execution continues -> the error was handled
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'P2' @(
        @{ Kind=1; Name='P2Case'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $wb = $book.Book; $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!P2_ResumeNext') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows  = @(Read-TraceRows $sx.ProcId)
    $exits = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    $r = @($exits | Where-Object { $_.function -eq 'P2_ResumeNext' })
    $outcome = if ($r.Count -and $r[0].outcome) { $r[0].outcome } else { '(no row)' }
    $wrote = [string]$wb.Worksheets.Item(1).Range('A1').Value2

    # The macro DID run to completion (A1 written), so the error was handled.
    Check 'macro-ran-past-the-error' ($wrote -eq 'after') "A1='$wrote' (statement after the swallowed error ran)"
    # By design. `outcome` reports errors that pass up the stack. An error raised and swallowed
    # inside one frame never crosses a frame boundary, and `handled` is reserved for a frame
    # that caught an error thrown below it. Reading Err.Number to relabel this would add a
    # hot-path COM read.
    Check 'same-frame-resume-next-reads-returned' ($outcome -eq 'returned') `
          "P2_ResumeNext outcome='$outcome' -- an error that never left its own frame must read 'returned'; 'handled' means a frame caught something thrown below it"
    # The one plainly-wrong reading: the code caught the error and carried on, so
    # it did NOT escape.
    Check 'does-not-read-threw' ($outcome -ne 'threw') "P2_ResumeNext outcome='$outcome' -- must not read as an escaped throw"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed -- outcome='$outcome'" }
    Complete-Test -Pass -Detail "same-frame Resume Next of a real error reads '$outcome' -- correct: the error never left its frame, so there is no chain to report"
}
catch { Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' })) }
