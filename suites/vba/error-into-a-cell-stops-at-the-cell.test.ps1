# AN ERROR THAT EXCEL TURNS INTO #VALUE! DOES NOT REACH THE MACRO THAT RECALCULATED.
#
# A macro calls Application.CalculateFull, and a cell's VBA function raises an error
# it does not handle. Excel puts #VALUE! in the cell, and the function reads
# unhandled; the macro carries on. The shadow stack still
# holds the macro beneath the function, so without a boundary the macro read
# `handled` for an error it never saw.
#
# A function that RETURNS an error value is different: it returned, and the value it
# returned is what the cell shows.
#
# A macro still running when disarm closes it has not thrown just because the raise
# opcode fired in it: object-model calls such as Application.Run fire it with no error.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
' Application.Run returns an error value for a macro it cannot run. Assigning it to
' a Double raises a type mismatch the function does not handle.
Public Function U_RunFails(ByVal x As Double) As Double
    Dim p As Double
    p = Application.Run("XRayNoSuchFunction", x)
    U_RunFails = p
End Function

Public Function U_Raises(ByVal x As Double) As Double
    Dim p As Double
    p = x
    Err.Raise 5, "XRayCase", "a deliberate error"
    U_Raises = p
End Function

' Andrew's cases: an error value RETURNED, with nothing raised.
Public Function GiveMeAHashValue()

GiveMeAHashValue = CVErr(xlErrValue)

End Function

Public Function SafeDiv(Numerator As Double, Denominator As Double) As Variant
    If Denominator = 0 Then
        SafeDiv = CVErr(xlErrDiv0)
    Else
        SafeDiv = Numerator / Denominator
    End If
End Function

' What Application.Run does for a macro it cannot find, as the cell reports it.
Public Function RunProbe() As String
    Dim v As Variant
    On Error GoTo Raised
    v = Application.Run("HGKJHKJ")
    RunProbe = "returned " & TypeName(v) & IIf(IsError(v), " (an error value)", "")
    Exit Function
Raised:
    RunProbe = "raised " & Err.Number
End Function

' Andrew's case: the error value cannot be assigned to a Double.
Public Function ABC() As Variant
    Dim p As Double
    p = Application.Run("HGKJHKJ")
    ABC = "Hello"
End Function

' Keeps Application.Run's error value in a Variant and returns it.
Public Function U_RunIntoVariant(ByVal x As Double) As Variant
    Dim v As Variant
    v = Application.Run("XRayNoSuchFunction", x)
    U_RunIntoVariant = v
End Function

Public Sub M_Recalc()
    Dim n As Long
    n = 1
    Application.CalculateFull
    n = 2
End Sub

Public Sub M_DisarmInside()
    Dim n As Long
    n = 1
    Application.Run "XRayXL_Disarm"
    n = 2
End Sub
'@

function ExitsOf($Rows, [string]$Fn) {
    @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}
function OutcomesOf($Rows, [string]$Fn) {
    $e = ExitsOf $Rows $Fn
    if ($e.Count -eq 0) { return '(no row)' }
    return (($e | ForEach-Object { "$($_.outcome)/$($_.trust)" }) -join ',')
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ErrIntoCell' @(
        @{ Kind=1; Name='ErrCell'; Code=$moduleCode }
    ) -Cells @{ 'A1' = '=U_RunFails(1)'; 'A2' = '=U_Raises(2)'; 'A3' = '=GiveMeAHashValue()'; 'A4' = '=U_RunIntoVariant(4)'; 'A5' = '=ABC()'; 'A6' = '=SafeDiv(1,0)'; 'A7' = '=SafeDiv(4,2)'; 'A8' = '=RunProbe()' }
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf
    $ws = $book.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    # On its own first, with no VBA frame beneath it.
    $ws.Range('A5').Formula = '=ABC()'
    $a5alone = Get-XRayCellText $ws.Range('A5')

    # Each run on its own: one that fails must not hide what the others do.
    $runErrors = @()
    foreach ($m in 'M_Recalc', 'M_DisarmInside') {
        try { $app.Run($leaf + '!' + $m) | Out-Null }
        catch { $runErrors += "$m -> $($_.Exception.Message)" }
        if ($m -eq 'M_Recalc') {
            $a1 = Get-XRayCellText $ws.Range('A1'); $a2 = Get-XRayCellText $ws.Range('A2')
            $a3 = Get-XRayCellText $ws.Range('A3'); $a4 = Get-XRayCellText $ws.Range('A4')
            $a6 = Get-XRayCellText $ws.Range('A6'); $a7 = Get-XRayCellText $ws.Range('A7')
            $a8 = Get-XRayCellText $ws.Range('A8')
        }
    }
    [void](Invoke-XRayDisarm $sx)
    Write-XRayObservation 'run-errors' ($(if ($runErrors) { $runErrors -join ' | ' } else { 'none' }))

    $rows = @(Read-TraceRows $sx.ProcId)

    Check 'excel-put-value-errors-in-the-cells' (($a1 -eq '#VALUE!') -and ($a2 -eq '#VALUE!')) "A1='$a1' A2='$a2'"

    $uRun = @(ExitsOf $rows 'U_RunFails'); $uRaise = @(ExitsOf $rows 'U_Raises')
    Check 'the-failing-functions-read-unhandled' `
          (($uRun.Count -ge 1) -and ($uRaise.Count -ge 1) -and
           (@(@($uRun) + @($uRaise) | Where-Object { $_.outcome -ne 'unhandled' }).Count -eq 0)) `
          "U_RunFails=$(OutcomesOf $rows 'U_RunFails') U_Raises=$(OutcomesOf $rows 'U_Raises')"

    $abc = @(ExitsOf $rows 'ABC')
    Check 'abc-reads-unhandled-alone-and-under-a-macro' `
          (($a5alone -eq '#VALUE!') -and ($abc.Count -ge 2) -and
           (@($abc | Where-Object { $_.outcome -ne 'unhandled' }).Count -eq 0)) `
          "A5='$a5alone' ABC=$(OutcomesOf $rows 'ABC')"

    # RETURNING AN ERROR IS NOT THROWING ONE: the row carries the value the cell shows.
    $rv = @(ExitsOf $rows 'GiveMeAHashValue')
    Check 'GiveMeAHashValue-returned-VALUE' `
          (($a3 -eq '#VALUE!') -and ($rv.Count -ge 1) -and
           (@($rv | Where-Object { $_.outcome -ne 'returned' -or $_.ret -ne '#VALUE!' }).Count -eq 0)) `
          "A3='$a3' GiveMeAHashValue=$(OutcomesOf $rows 'GiveMeAHashValue') ret='$(@($rv | ForEach-Object { $_.ret }) -join ',')'"

    $sd = @(ExitsOf $rows 'SafeDiv')
    $sdRets = @($sd | ForEach-Object { $_.ret })
    Check 'SafeDiv-returned-DIV0-and-a-number' `
          (($a6 -eq '#DIV/0!') -and ($a7 -eq '2') -and ($sd.Count -ge 2) -and
           (@($sd | Where-Object { $_.outcome -ne 'returned' }).Count -eq 0) -and
           ($sdRets -contains '#DIV/0!') -and ($sdRets -contains '2')) `
          "A6='$a6' A7='$a7' SafeDiv=$(OutcomesOf $rows 'SafeDiv') ret='$($sdRets -join ',')'"

    # Observed, not asserted: what Application.Run does for a macro it cannot find.
    Write-XRayObservation 'application-run-into-a-variant' "A4='$a4' U_RunIntoVariant=$(OutcomesOf $rows 'U_RunIntoVariant')"
    Write-XRayObservation 'application-run-probe' "A8='$a8'"

    $recalc = @(ExitsOf $rows 'M_Recalc')
    Check 'the-recalculating-macro-returned' `
          (($recalc.Count -eq 1) -and ($recalc[0].outcome -eq 'returned')) `
          "M_Recalc=$(OutcomesOf $rows 'M_Recalc')"

    # NOTHING NESTS UNDER A DEAD FRAME. Excel calculates cells one after another, so a
    # cell's function sits directly under the macro, or at the top, never under another
    # cell's function; and a macro started from outside sits at the top.
    $recalcSpan = if ($recalc.Count -eq 1) { [string]$recalc[0].span } else { '' }
    $entries = @($rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })
    $cellFns = @('U_RunFails', 'U_Raises', 'GiveMeAHashValue', 'U_RunIntoVariant', 'ABC', 'SafeDiv', 'RunProbe')
    $misnested = @($entries | Where-Object {
        ($cellFns -contains $_.function -and $_.parent -ne '0' -and $_.parent -ne $recalcSpan) -or
        ($_.function -like 'M_*' -and $_.parent -ne '0') })
    Check 'nothing-nests-under-a-dead-frame' ($misnested.Count -eq 0) `
          ("misnested: " + (@($misnested | ForEach-Object { "$($_.function)(parent $($_.parent), depth $($_.depth))" }) -join ', '))

    $d = @(ExitsOf $rows 'M_DisarmInside')
    Check 'a-macro-open-at-disarm-did-not-throw' `
          (($d.Count -eq 1) -and ($d[0].outcome -eq 'returned') -and ($d[0].trust -eq 'flush')) `
          "M_DisarmInside=$(OutcomesOf $rows 'M_DisarmInside')"


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'an unhandled error reads unhandled and stops at the cell; a returned error is returned; disarm is not a throw'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
