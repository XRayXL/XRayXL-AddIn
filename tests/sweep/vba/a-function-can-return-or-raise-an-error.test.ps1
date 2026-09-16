# A FUNCTION THAT RETURNS AN ERROR VALUE vs ONE THAT RAISES AN ERROR.
#
# Two things a worksheet cell can show as `#...`, and the trace must tell them
# apart:
#
#   RETURNS an error   CVErr(xlErr...) -- the function finished. `outcome` is
#                      `returned` and `ret` is the error, spelt as Excel spells it.
#   RAISES an error    Err.Raise with no handler -- the function did not finish.
#                      Excel shows `#VALUE!` and `outcome` is `unhandled`.
#
# The first is exercised for every Excel error a VBA function can hand back, so a
# decoder that mixed two of them up, or blanked one, is caught. Nothing here is
# raised except the one function that means to, so no cell but that one is
# `unhandled`.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# name -> @(cell formula, the text Excel and the trace should both show)
$errCases = [ordered]@{
    'RetNull'  = @('=RetNull()',  '#NULL!')
    'RetDiv0'  = @('=RetDiv0()',  '#DIV/0!')
    'RetValue' = @('=RetValue()', '#VALUE!')
    'RetRef'   = @('=RetRef()',   '#REF!')
    'RetName'  = @('=RetName()',  '#NAME?')
    'RetNum'   = @('=RetNum()',   '#NUM!')
    'RetNA'    = @('=RetNA()',    '#N/A')
}

$moduleCode = @'
' Each returns one Excel error value. Nothing is raised: the function returns.
Public Function RetNull() As Variant
    RetNull = CVErr(xlErrNull)
End Function
Public Function RetDiv0() As Variant
    RetDiv0 = CVErr(xlErrDiv0)
End Function
Public Function RetValue() As Variant
    RetValue = CVErr(xlErrValue)
End Function
Public Function RetRef() As Variant
    RetRef = CVErr(xlErrRef)
End Function
Public Function RetName() As Variant
    RetName = CVErr(xlErrName)
End Function
Public Function RetNum() As Variant
    RetNum = CVErr(xlErrNum)
End Function
Public Function RetNA() As Variant
    RetNA = CVErr(xlErrNA)
End Function

' One line, one raise, no handler: Excel turns it into #VALUE! in the cell.
Public Function RaiseObjErr() As Variant
    Err.Raise vbObjectError
End Function
'@

function ExitsOf($Rows, [string]$Fn) {
    @($Rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' -and $_.function -eq $Fn })
}

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $cells = @{}
    $row = 1
    foreach ($name in $errCases.Keys) { $cells["A$row"] = $errCases[$name][0]; $row++ }
    $cells["A$row"] = '=RaiseObjErr()'; $raiseRow = $row

    New-XRayMacroBook $sx 'RetOrRaise' @(
        @{ Kind=1; Name='RetOrRaise'; Code=$moduleCode }
    ) -Cells $cells
    $book = Get-XRayMacroBook
    $ws = $book.Sheet

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
    if ($armLine -match 'NO ERROR ATTRIBUTION') {
        Complete-Test -Fail -Detail "the raise slot did not verify on this VBE7: $armLine"
    }

    $app.CalculateFull()

    # Read the cells while still armed, then disarm to flush the trace.
    $cellText = @{}
    $row = 1
    foreach ($name in $errCases.Keys) { $cellText[$name] = Get-XRayCellText $ws.Range("A$row"); $row++ }
    $raiseCell = Get-XRayCellText $ws.Range("A$raiseRow")

    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)

    # ---- every RETURNED error: outcome returned, ret is that error --------------
    foreach ($name in $errCases.Keys) {
        $want = $errCases[$name][1]
        $ex = @(ExitsOf $rows $name)
        $cellOk = $cellText[$name] -eq $want
        $haveExit = $ex.Count -ge 1
        $allReturned = (@($ex | Where-Object { $_.outcome -ne 'returned' -or $_.ret -ne $want }).Count -eq 0)
        Check "returned-$name" ($cellOk -and $haveExit -and $allReturned) `
              ("cell='{0}' want='{1}' outcome='{2}' ret='{3}'" -f $cellText[$name], $want,
               (@($ex | ForEach-Object { $_.outcome }) -join ','),
               (@($ex | ForEach-Object { $_.ret }) -join ','))
    }

    # ---- the RAISED error: cell #VALUE!, outcome unhandled ----------------------
    $rex = @(ExitsOf $rows 'RaiseObjErr')
    $raiseCellOk = $raiseCell -eq '#VALUE!'
    $haveRaise = $rex.Count -ge 1
    $allUnhandled = (@($rex | Where-Object { $_.outcome -ne 'unhandled' }).Count -eq 0)
    Check 'raised-vbObjectError-is-unhandled' ($raiseCellOk -and $haveRaise -and $allUnhandled) `
          ("cell='{0}' outcome='{1}'" -f $raiseCell, (@($rex | ForEach-Object { $_.outcome }) -join ','))

    # ---- the negative control: only the raise is unhandled ----------------------
    $vbaExits = @($rows | Where-Object { $_.kind -eq 'exit' -and $_.source -eq 'VBA' })
    $stray = @($vbaExits | Where-Object { $_.outcome -eq 'unhandled' -and $_.function -ne 'RaiseObjErr' })
    Check 'only-the-raise-is-unhandled' ($stray.Count -eq 0) `
          ("also unhandled: " + (@($stray | ForEach-Object { $_.function }) -join ','))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'every returned error reads returned with its value; the one raise reads unhandled'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
