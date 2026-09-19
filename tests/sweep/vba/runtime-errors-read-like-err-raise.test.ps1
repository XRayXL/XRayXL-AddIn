# An error the VBA runtime raises reads the same as one Err.Raise raises.
#
# Division by zero, an overflow and a failed conversion are raised inside the opcode that did the
# arithmetic, not through Err.Raise. The frame still leaves without its epilogue, and that is what
# makes it `threw`: its caller that catches reads `handled`, one that passes it on `unwound`, and a
# worksheet function it escapes from `unhandled`.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Private Function H_Div(ByVal a As Double, ByVal b As Double) As Double
    H_Div = a / b
End Function
Private Function H_Overflow(ByVal a As Double) As Integer
    H_Overflow = CInt(a)
End Function
Private Function H_Mismatch(ByVal s As String) As Double
    H_Mismatch = CDbl(s)
End Function
Private Function H_DivRaw(ByVal a As Double, ByVal b As Double) As Double
    H_DivRaw = a / b
End Function
Private Function H_Clean(ByVal a As Double, ByVal b As Double) As Double
    H_Clean = a / b
End Function

Public Function U_DivCaught(ByVal b As Double) As Variant
    On Error GoTo F
    U_DivCaught = H_Div(1, b): Exit Function
F:  U_DivCaught = "caught"
End Function
Public Function U_OverflowCaught(ByVal a As Double) As Variant
    On Error GoTo F
    U_OverflowCaught = H_Overflow(a): Exit Function
F:  U_OverflowCaught = "caught"
End Function
Public Function U_MismatchCaught(ByVal s As String) As Variant
    On Error GoTo F
    U_MismatchCaught = H_Mismatch(s): Exit Function
F:  U_MismatchCaught = "caught"
End Function
' Nothing catches it: Excel shows #VALUE!.
Public Function U_DivRaw(ByVal b As Double) As Double
    U_DivRaw = H_DivRaw(1, b)
End Function
Public Function U_Clean(ByVal b As Double) As Variant
    On Error GoTo F
    U_Clean = H_Clean(1, b): Exit Function
F:  U_Clean = "caught"
End Function

' The chain again, from a macro: R_Thrower divides by zero, R_Middle passes it on, R_Outer catches.
Public Sub R_Outer()
    Dim n As Long
    On Error GoTo Caught
    n = 1
    R_Middle
    n = 2
    Exit Sub
Caught:
    n = 3
    n = 4
End Sub
Public Sub R_Middle()
    Dim m As Long
    m = 1
    R_Thrower 0
    m = 2
End Sub
Public Sub R_Thrower(ByVal zero As Double)
    Dim t As Double
    t = 1 / zero
    t = 2
End Sub
'@

function OutcomeOf($Rows, [string]$Fn) {
    $e = @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
    if ($e.Count -eq 0) { return '(no row)' }
    return (($e | ForEach-Object { $_.outcome }) -join ',')
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'RuntimeErr' @(
        @{ Kind=1; Name='RtErr'; Code=$moduleCode }
    ) -Cells @{ 'A1' = '=U_DivCaught(0)'; 'A2' = '=U_OverflowCaught(40000)'; 'A3' = '=U_MismatchCaught("x")';
                'A4' = '=U_DivRaw(0)'; 'A5' = '=U_Clean(2)' }
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf
    $ws = $book.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.CalculateFull()
    $cells = @(1..5 | ForEach-Object { Get-XRayCellText $ws.Range("A$_") })
    $app.Run($leaf + '!R_Outer') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    Check 'excel-saw-what-vba-did' `
          (($cells[0] -eq 'caught') -and ($cells[1] -eq 'caught') -and ($cells[2] -eq 'caught') -and
           ($cells[3] -eq '#VALUE!') -and ($cells[4] -eq '0.5')) "A1:A5 = $($cells -join ' | ')"

    foreach ($pair in @(@('H_Div', 'U_DivCaught'), @('H_Overflow', 'U_OverflowCaught'), @('H_Mismatch', 'U_MismatchCaught'))) {
        $h = OutcomeOf $rows $pair[0]; $u = OutcomeOf $rows $pair[1]
        Check "$($pair[0])-threw-and-$($pair[1])-handled" (($h -eq 'threw') -and ($u -eq 'handled')) "$($pair[0])='$h' $($pair[1])='$u'"
    }

    $h = OutcomeOf $rows 'H_DivRaw'; $u = OutcomeOf $rows 'U_DivRaw'
    Check 'an-uncaught-runtime-error-reads-unhandled-at-the-cell' (($h -eq 'threw') -and ($u -eq 'unhandled')) "H_DivRaw='$h' U_DivRaw='$u'"

    $t = OutcomeOf $rows 'R_Thrower'; $m = OutcomeOf $rows 'R_Middle'; $o = OutcomeOf $rows 'R_Outer'
    Check 'a-macro-chain-reads-threw-unwound-handled' (($t -eq 'threw') -and ($m -eq 'unwound') -and ($o -eq 'handled')) `
          "R_Thrower='$t' R_Middle='$m' R_Outer='$o'"

    # THE NEGATIVE CONTROL: the same shape with nothing raised.
    $h = OutcomeOf $rows 'H_Clean'; $u = OutcomeOf $rows 'U_Clean'
    Check 'the-clean-pair-returned' (($h -eq 'returned') -and ($u -eq 'returned')) "H_Clean='$h' U_Clean='$u'"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'division by zero, overflow and a failed conversion read threw / handled / unwound / unhandled'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
