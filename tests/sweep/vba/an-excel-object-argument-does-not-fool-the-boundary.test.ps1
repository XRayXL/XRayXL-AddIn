# An Excel object passed to a VBA method must not look like Excel starting the frame.
#
# A Sub with a handler calls a class method, passing Range("A1"), and the method raises. The
# error must cross back to the Sub: the method reads `threw` and the Sub `handled`, never
# `unhandled`.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$classCode = @'
' Takes an Excel object, then raises. No handler here: the error must leave.
Public Sub TouchAndRaise(ByVal r As Object)
    Dim n As Long
    n = r.Row
    Err.Raise 5, "XRayCase", "raised after receiving a Range"
End Sub
'@

$moduleCode = @'
Public Sub O_Outer()
    Dim o As Object
    On Error GoTo Caught
    Set o = New XRArg
    o.TouchAndRaise ThisWorkbook.Worksheets(1).Range("A1")
    Exit Sub
Caught:
    Dim n As Long
    n = 3
    n = 4
End Sub
'@

function ExitsOf($Rows, [string]$Fn) {
    @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}
function OutcomesOf($Rows, [string]$Fn) {
    $e = ExitsOf $Rows $Fn
    if ($e.Count -eq 0) { return '(no row)' }
    return (($e | ForEach-Object { $_.outcome }) -join ',')
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ObjArg' @(
        @{ Kind=1; Name='ObjArg';  Code=$moduleCode }
        @{ Kind=2; Name='XRArg';   Code=$classCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!O_Outer') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    $m = @(ExitsOf $rows 'TouchAndRaise')
    $o = @(ExitsOf $rows 'O_Outer')

    Check 'the-method-error-crossed-back-to-vba' `
          (($m.Count -ge 1) -and (@($m | Where-Object { $_.outcome -ne 'threw' }).Count -eq 0)) `
          "TouchAndRaise=$(OutcomesOf $rows 'TouchAndRaise') -- must be threw, not unhandled"
    Check 'the-vba-handler-caught-it' `
          (($o.Count -ge 1) -and (@($o | Where-Object { $_.outcome -ne 'handled' }).Count -eq 0)) `
          "O_Outer=$(OutcomesOf $rows 'O_Outer')"
    Check 'nothing-here-reads-unhandled' `
          (@($rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.outcome -eq 'unhandled' }).Count -eq 0) `
          "the Range argument must not read as Excel starting the frame"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'an Excel object argument did not fool the boundary; the error chain held'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
