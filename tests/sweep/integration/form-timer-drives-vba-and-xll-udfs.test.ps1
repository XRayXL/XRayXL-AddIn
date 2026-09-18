# A UserForm with a timer, end to end: VBA tracing is not confined to UDFs reached from a
# recalc, and form code itself appears.
#
# A standard-module Sub instantiates a UserForm and calls its method Arm, which schedules
# Application.OnTime. The handler XR_TimerTick dirties the UDF inputs and calls
# Worksheet.Calculate on a sheet holding one VBA UDF and one XLL UDF. One trace file must hold:
#
#   UserForm_Initialize   FORM EVENT CODE -- the interpreter invokes it as part
#                         of loading the form, not via a call opcode. THIS is the
#                         row that proves "form VBA is in the trace".
#   Arm                   a FORM METHOD, called through the form instance
#   XR_TimerTick          a plain Sub entered by the TIMER (OnTime), not a cell
#                         and not Application.Run -- general, event-driven VBA
#   XR_TimerUdf           a VBA UDF, called from cell A1 during the recalc
#   XR_TimerHelper        called BY the UDF -- VBA nesting inside a UDF
#   TxB                   the XLL UDF in cell B1, source=XLL
#
# OnTime, not a shown form with a timer control: a shown form needs a pumped message loop, which
# a COM driver cannot reliably provide. The form is never .Show'd; Initialize fires on first
# member access.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public Sub XR_StartForm()
    Dim f As New XRTimerForm       ' UserForm_Initialize fires on first use below
    f.Arm                          ' a FORM METHOD; it schedules the timer
    Set f = Nothing
End Sub

Public Sub XR_TimerTick()
    ' The scheduled handler -- the "timer". Reads the worksheet, dirties the
    ' UDF inputs, recalculates. This is VBA entered by neither a cell nor
    ' Application.Run.
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("S1")
    Dim used As Range
    Set used = ws.UsedRange                                ' read the worksheet
    ws.Range("C1").Value = ws.Range("C1").Value + 1        ' dirty A1 and B1
    ws.Calculate                                           ' run the two UDFs
    ws.Range("Z1").Value = "done"                          ' completion latch
End Sub

Public Function XR_TimerUdf(ByVal a As Long, ByVal c As Long) As Long
    XR_TimerUdf = XR_TimerHelper(a) + c                    ' VBA UDF -> nested VBA
End Function

Public Function XR_TimerHelper(ByVal n As Long) As Long
    XR_TimerHelper = n * 2
End Function
'@

$formCode = @'
Private Sub UserForm_Initialize()
    Dim marker As Long
    marker = 1                      ' a statement, so the interpreter has a frame
End Sub

Public Sub Arm()
    ' The form owns the timer: fire XR_TimerTick as soon as Excel idles.
    Application.OnTime Now, ThisWorkbook.Name & "!XR_TimerTick"
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'FormTimer' @(
        @{ Kind=1; Name='TimerCase';   Code=$moduleCode }
        @{ Kind=3; Name='XRTimerForm'; Code=$formCode }
    ) @{
        'A1' = '=XR_TimerUdf(2,C1)'   # VBA UDF, depends on C1
        'B1' = '=TxB(4,C1)'           # XLL UDF, depends on C1
    } {
        param($ws)
        $ws.Range('C1').Value = 5     # what both formulas depend on
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $leaf = $book.Leaf

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    # Reset the completion latch, kick the form, then let Excel idle so OnTime
    # fires. Poll the latch rather than guess a sleep -- OnTime lands on the
    # next idle, and a light cell read between sleeps is how Excel gets it.
    $ws.Range('Z1').Value = ''
    $app.Run($leaf + '!XR_StartForm') | Out-Null
    $fired = Wait-XRayCondition { ([string]$ws.Range('Z1').Value2) -eq 'done' } 12 250
    # the handler writes Z1 before the recalc it starts has settled
    [void](Wait-XRayCalcDone $app)
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    Check 'timer-fired' $fired "OnTime handler ran (Z1='$([string]$ws.Range('Z1').Value2)')"

    $rows    = Select-BookRows (Read-TraceRows $sx.ProcId) $leaf
    $vEntry  = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $xEntry  = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'XLL') })
    $vNames  = @($vEntry | ForEach-Object { $_.function })
    $xNames  = @($xEntry | ForEach-Object { $_.function })

    # ---- FORM EVENT CODE is in the trace (the headline claim) -------------
    Check 'form-initialize-traced-as-vba' ($vNames -contains 'UserForm_Initialize') `
          ("VBA entries: " + ($vNames -join ','))

    # ---- the form's own METHOD ran and was traced ------------------------
    Check 'form-method-traced-as-vba' ($vNames -contains 'Arm') `
          ("VBA entries: " + ($vNames -join ','))

    # ---- the TIMER handler -- VBA entered by neither a cell nor Run -------
    Check 'timer-handler-traced-as-vba' ($vNames -contains 'XR_TimerTick') `
          ("VBA entries: " + ($vNames -join ','))

    # ---- the VBA UDF, from a cell ----------------------------------------
    $udf = @($vEntry | Where-Object { $_.function -eq 'XR_TimerUdf' })
    Check 'vba-udf-traced-from-cell' (($udf.Count -ge 1) -and ($udf[0].caller -eq 'cell')) `
          ("XR_TimerUdf count=$($udf.Count) caller='$(if ($udf.Count) { $udf[0].caller } else { '' })'")

    # ---- the VBA UDF's nested call ---------------------------------------
    $helper = @($vEntry | Where-Object { $_.function -eq 'XR_TimerHelper' })
    $udfSpan = if ($udf.Count) { [string]$udf[0].span } else { '' }
    $helperParent = if ($helper.Count) { [string]$helper[0].parent } else { '' }
    Check 'vba-udf-nests-its-helper' (($helper.Count -ge 1) -and ($helperParent -eq $udfSpan) -and ($udfSpan -ne '')) `
          ("XR_TimerHelper parent=$helperParent, XR_TimerUdf span=$udfSpan")

    # ---- the XLL UDF, in the SAME trace file, source=XLL ------------------
    $txb = @($xEntry | Where-Object { $_.function -eq 'TxB' })
    Check 'xll-udf-traced-from-cell' (($txb.Count -ge 1) -and ($txb[0].caller -eq 'cell')) `
          ("TxB count=$($txb.Count) caller='$(if ($txb.Count) { $txb[0].caller } else { '' })'  (XLL entries: " + ($xNames -join ',') + ")")

    # ---- named, not addresses --------------------------------------------
    $unnamed = @($vEntry | Where-Object { $_.function -match '^0x[0-9A-Fa-f]+$' })
    Check 'no-vba-procedure-reported-as-an-address' ($unnamed.Count -eq 0) `
          ("unnamed: " + (@($unnamed | ForEach-Object { $_.function }) -join ','))

    # Nothing errored, so nothing may read as though it did. Ordinary object-model VBA (a cell
    # write, Application.OnTime, .Calculate) reaches the same raise opcode (497) as Err.Raise,
    # and a raiser that runs its epilogue did not throw.
    $vExit  = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    $notRet = @($vExit | Where-Object { $_.outcome -ne 'returned' })
    Check 'no-benign-object-model-raise-reads-as-an-error' ($notRet.Count -eq 0) `
          ("frames not returned: " + (@($notRet | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    # ---- caller invariants hold across VBA and XLL together --------------
    $ci = Test-RowInvariants $rows
    Check 'caller-invariants-hold' ($ci.Count -eq 0) (($ci -join '; '))

    Write-Output ''
    Write-Output 'VBA entries:'
    foreach ($e in $vEntry) { Write-Output ("  {0,-20} caller={1,-6} depth={2} parent={3}" -f $e.function, $e.caller, $e.depth, $e.parent) }
    Write-Output 'XLL entries:'
    foreach ($e in $xEntry) { Write-Output ("  {0,-20} caller={1}" -f $e.function, $e.caller) }

    [void](Set-XRayTraceParam $sx 'VBA' 'DEPTH' 'OFF')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "form + timer captured: form code, timer handler, VBA UDF (+helper) and XLL UDF all in one trace"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
