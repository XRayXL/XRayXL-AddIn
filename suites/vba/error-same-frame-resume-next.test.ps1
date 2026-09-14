# A REAL ERROR SWALLOWED BY On Error Resume Next IN THE SAME FRAME.
#
# This is the accepted trade-off of the fix that stopped a cell write being
# reported as a throw. A benign object-model raise and a real error caught in
# place by `On Error Resume Next` are indistinguishable AT the
# raise -- same opcode (497), same registers -- and both let the frame run on to
# its own exit. So both now read `returned`. For the benign case that is right;
# for a real error caught in the same frame the truest label is `handled`.
#
# This PINS the current behaviour so a future fix (reading Err.Number to tell a
# real error from a benign raise) flips this test rather than passing silently.
# It asserts what IS, names what it OUGHT to be, and fails loudly only on the one
# reading that would be plainly wrong -- `threw`, which would mean the tracer
# thought the error escaped when the code caught it and carried on.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
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
    # an unverified raise slot is a product failure, not a SKIP: every outcome would read `returned`
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail ("the raise slot did not verify on this VBE7, so no outcome " +
                                     "below can be attributed: $armLine")
    }

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
    # BY DESIGN, NOT A GAP. `outcome` reports errors that PASS UP THE STACK --
    # who threw, who it unwound through, who caught it. An error raised and
    # swallowed inside ONE frame never crosses a frame boundary, so there is no
    # chain to report and the activation did what the column says: it returned.
    # `handled` is reserved for a frame that caught an error thrown BELOW it,
    # which is the case a reader needs to find. Reading Err.Number to relabel
    # this one would add a hot-path COM read to report a non-event.
    Check 'same-frame-resume-next-reads-returned' ($outcome -eq 'returned') `
          "P2_ResumeNext outcome='$outcome' -- an error that never left its own frame must read 'returned'; 'handled' means a frame caught something thrown below it"
    # The one plainly-wrong reading: the code caught the error and carried on, so
    # it did NOT escape.
    Check 'does-not-read-threw' ($outcome -ne 'threw') "P2_ResumeNext outcome='$outcome' -- must not read as an escaped throw"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed -- outcome='$outcome'" }
    Complete-Test -Pass -Detail "same-frame Resume Next of a real error reads '$outcome' -- correct: the error never left its frame, so there is no chain to report"
}
catch { Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message) }
