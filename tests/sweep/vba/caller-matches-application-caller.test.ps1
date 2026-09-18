# `caller` reproduces Application.Caller. That is the whole specification.
#
# The tracer asks Excel12(xlfCaller) once per activation, at frame-open, so a row says what
# Application.Caller would have returned on the first line of the procedure. This test inserts
# that call and compares, rather than predicting Excel's answer: the documented table says
# nothing about nesting or event handlers.
#
# Each procedure reads Application.Caller in its own body, not in a helper, since whether the
# VBA call stack affects the answer is part of the question.
#
# Not covered: a macro on a button. Application.Caller names the shape only for a real mouse
# click, which this project will not simulate.
#
# The caller is inherited, not reset per activation: a Worksheet_Change nested inside Auto_Open
# reports Auto_Open's caller. Case 4 cannot show this, because its outer and inner callers are
# both #REF!; the verdict is drawn from case 5.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# The formatter mirrors src/xll/caller.cpp's vocabulary exactly, because the
# whole test is a string comparison against what that file wrote.
$moduleCode = @'
Public gLog As String

Public Sub XRReset()
    gLog = ""
End Sub

Public Function XRGetLog() As String
    XRGetLog = gLog
End Function

' v is captured INLINE by the caller and passed in; nothing here reads
' Application.Caller.
Public Sub XRNote(ByVal proc As String, ByVal v As Variant)
    Dim s As String
    If IsObject(v) Then
        s = "cell"
        On Error Resume Next
        s = "cell|" & v.Address(False, False) & "|" & v.Parent.Name
        If Err.Number <> 0 Then s = "cell|?|?": Err.Clear
        On Error GoTo 0
    ElseIf IsArray(v) Then
        s = "array"
    ElseIf IsError(v) Then
        Dim n As Long
        n = Val(Mid$(CStr(v), 7))       ' "Error 2023" -> 2023
        Select Case n
            Case 2000: s = "none:null"
            Case 2007: s = "none:div0"
            Case 2015: s = "none:value"
            Case 2023: s = "none:ref"
            Case 2029: s = "none:name"
            Case 2036: s = "none:num"
            Case 2042: s = "none:na"
            Case Else: s = "none:err" & CStr(n)
        End Select
    Else
        s = "object:" & CStr(v)
    End If
    gLog = gLog & proc & "=" & s & ";"
End Sub

' ---- the procedures under test. Each reads Application.Caller INLINE. -------

Public Function XRUdf(ByVal x As Double) As Double
    Dim v As Variant
    If IsObject(Application.Caller) Then Set v = Application.Caller Else v = Application.Caller
    XRNote "XRUdf", v
    XRUdf = x * 2
End Function

Public Sub XRFromRun()
    Dim v As Variant
    If IsObject(Application.Caller) Then Set v = Application.Caller Else v = Application.Caller
    XRNote "XRFromRun", v
End Sub

' A DISTINCTIVE OUTER CALLER, WITHOUT A CLICK. Microsoft documents Auto_Open as
' returning the NAME OF THE DOCUMENT as text -- unlike Application.Run, whose
' answer (#REF!) is the same one an event handler gives, so a Run-driven nesting
' cannot tell "inherited" from "reset". This one can, if the caller context
' survives RunAutoMacros.
Public Sub Auto_Open()
    Dim v As Variant
    If IsObject(Application.Caller) Then Set v = Application.Caller Else v = Application.Caller
    XRNote "Auto_Open", v
    Application.EnableEvents = True
    ThisWorkbook.Worksheets("S1").Range("B3").Value = _
        ThisWorkbook.Worksheets("S1").Range("B3").Value + 1
End Sub

' THE NESTED CASE. Invoked by Application.Run, writes a cell with events on, so
' Worksheet_Change fires INSIDE this activation. Both record.
Public Sub XRWriteCell()
    Dim v As Variant
    If IsObject(Application.Caller) Then Set v = Application.Caller Else v = Application.Caller
    XRNote "XRWriteCell", v
    Application.EnableEvents = True
    ThisWorkbook.Worksheets("S1").Range("B2").Value = _
        ThisWorkbook.Worksheets("S1").Range("B2").Value + 1
End Sub
'@

$sheetCode = @'
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim v As Variant
    If IsObject(Application.Caller) Then Set v = Application.Caller Else v = Application.Caller
    XRNote "Worksheet_Change", v
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # The sheet's own code module, reached by the sheet's CODE NAME.
    New-XRayMacroBook $sx 'Caller' @(
        @{ Kind=1; Name='XRCase'; Code=$moduleCode }
        @{ Kind='Sheet'; Code=$sheetCode }
    ) @{} {
        param($ws)
        $ws.Range('B2').Value = 0
    }
    $book = Get-XRayMacroBook
    $wb = $book.Book; $ws = $book.Sheet; $leaf = $book.Leaf
    $bang = "$leaf!"

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    # One arming session per trigger, so each trigger's rows stand alone --
    # csv::Open truncates at every arm.
    function Invoke-Trigger([scriptblock]$act) {
        $app.Run($bang + 'XRReset') | Out-Null
        $mark = Get-LogLength $paths.Log
        [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
        $line = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
        if ($line -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $line" }
        & $act
        [void](Wait-XRayCalcDone $app)
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }
        $vba = [string]$app.Run($bang + 'XRGetLog')
        $rows = @(Read-TraceRows $sx.ProcId | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
        return [pscustomobject]@{ Vba = $vba; Rows = $rows }
    }

    # What the tracer wrote for a procedure, in the VBA formatter's vocabulary.
    function Get-TracerCaller($rows, [string]$proc) {
        $r = @($rows | Where-Object { $_.function -eq $proc })
        if ($r.Count -eq 0) { return $null }
        $x = $r[0]
        # Translated into the VBA formatter's vocabulary here, so the VBA side stays an
        # independent witness. It says `object:` for any string (the trace's kind is `name`) and
        # a bare `array` for anything IsArray covers, which includes a toolbar and a menu.
        if ($x.caller -eq 'cell') { return "cell|$(Get-CallerCell $x)|$(Get-CallerSheet $x)" }
        if ($x.caller -eq 'name') { return "object:$($x.callerref)" }
        if ($x.caller -in @('toolbar','menu','array')) { return 'array' }
        if ($x.caller -eq 'none') { return "none:$($x.callerref)" }
        return [string]$x.caller
    }
    function Get-VbaCaller([string]$log, [string]$proc) {
        foreach ($part in $log.Split(';')) {
            if ($part -like "$proc=*") { return $part.Substring($proc.Length + 1) }
        }
        return $null
    }
    # The tracer writes the sheet as [Book]Sheet; VBA's Parent.Name is the bare
    # sheet. Compare on the parts that mean the same thing.
    function Same([string]$tracer, [string]$vba) {
        if ($null -eq $tracer -or $null -eq $vba) { return $false }
        if ($tracer -eq $vba) { return $true }
        if ($tracer -like 'cell|*' -and $vba -like 'cell|*') {
            $t = $tracer.Split('|'); $v = $vba.Split('|')
            return ($t[1] -eq $v[1]) -and ($t[2] -like "*$($v[2])*")
        }
        return $false
    }

    # Results go to a script variable, not down the pipeline: a function's return value is
    # everything it wrote to the pipeline, so an assignment would swallow the `STRETCH case=`
    # lines Check emits.
    $script:lastCmp = $null
    function Compare-Proc($res, [string]$proc, [string]$label) {
        $t = Get-TracerCaller $res.Rows $proc
        $v = Get-VbaCaller $res.Vba $proc
        $script:lastCmp = @{ Tracer = $t; Vba = $v }
        Check $label (Same $t $v) "tracer='$t' vba='$v'   (vba log: $($res.Vba))"
    }

    # ---- 1. a UDF in a cell: the documented Range case ---------------------
    $r1 = Invoke-Trigger { $ws.Range('A1').Formula = '=XRUdf(21)'; $app.CalculateFullRebuild() }
    Compare-Proc $r1 'XRUdf' 'udf-in-a-cell-agrees'

    # ---- 2. Application.Run: no caller on a sheet --------------------------
    $r2 = Invoke-Trigger { $app.Run($bang + 'XRFromRun') | Out-Null }
    Compare-Proc $r2 'XRFromRun' 'application-run-agrees'

    # ---- 3. an edit firing Worksheet_Change --------------------------------
    $r3 = Invoke-Trigger {
        $app.EnableEvents = $true
        $ws.Range('B2').Value = 5
        $app.EnableEvents = $false
    }
    Compare-Proc $r3 'Worksheet_Change' 'edit-fires-change-and-agrees'
    $e = $script:lastCmp

    # ---- 4. THE NESTED CASE ------------------------------------------------
    # Application.Run -> XRWriteCell -> writes a cell -> Worksheet_Change, one
    # activation inside another. Whatever Excel answers in the inner frame, the
    # tracer must answer the same.
    $r4 = Invoke-Trigger {
        $app.EnableEvents = $true
        $app.Run($bang + 'XRWriteCell') | Out-Null
        $app.EnableEvents = $false
    }
    Compare-Proc $r4 'XRWriteCell'      'nested-outer-agrees'
    $outer = $script:lastCmp
    Compare-Proc $r4 'Worksheet_Change' 'nested-inner-agrees'
    $inner = $script:lastCmp

    # Both frames must be present, or the comparison above passed vacuously on
    # a nesting that never happened.
    $both = ($null -ne $outer.Tracer) -and ($null -ne $inner.Tracer)
    Check 'nested-both-frames-traced' $both "outer='$($outer.Tracer)' inner='$($inner.Tracer)'"

    # 5. The discriminating nesting. Application.Run and an event handler both answer #REF!, so
    # case 4 cannot separate "inherited" from "reset". Auto_Open is documented to answer with
    # the document name, so the nested Change settles it either way.
    #
    # RunAutoMacros, because Workbooks.Open from automation does not run Auto_Open. A null
    # answer here is a result, not a failure: the assertion is only that the tracer agrees with
    # the VBA.
    $r5 = Invoke-Trigger {
        $app.EnableEvents = $true
        try { $wb.RunAutoMacros(1) } catch { }      # 1 = xlAutoOpen
        $app.EnableEvents = $false
    }
    Compare-Proc $r5 'Auto_Open' 'auto-open-agrees'
    $ao = $script:lastCmp
    Compare-Proc $r5 'Worksheet_Change' 'auto-open-nested-inner-agrees'
    $aoInner = $script:lastCmp

    $discriminated = ($null -ne $ao.Vba) -and ($ao.Vba -notlike 'none:*')
    if ($discriminated) {
        $aoDetail = "yes: Auto_Open caller='$($ao.Vba)', inner Change caller='$($aoInner.Vba)' -- nesting is decidable from this case"
    } else {
        $aoDetail = "no: Auto_Open caller='$($ao.Vba)' via RunAutoMacros, so it cannot separate inherited from reset either"
    }
    Write-XRayObservation 'auto-open-is-a-distinctive-outer-caller' $aoDetail

    # What Excel does, recorded either way; not an assertion. The verdict comes from case 5:
    # only a distinctive outer caller can tell the two hypotheses apart.
    if ($discriminated) {
        $verdict = if ($aoInner.Vba -eq $ao.Vba) {
            "INHERITED -- the nested handler reports the OUTER invocation's caller ('$($ao.Vba)')"
        } else {
            "RESET -- outer='$($ao.Vba)' but the nested handler says '$($aoInner.Vba)'"
        }
    } else {
        $verdict = 'UNDECIDED -- no case produced a distinctive outer caller'
    }
    Write-XRayObservation 'observed-semantics' (
        "edit->Change caller='{0}'; Run->WriteCell caller='{1}'; its nested Change caller='{2}' (both #REF!, so vacuous); VERDICT from Auto_Open: {3}" -f
        $e.Vba, $outer.Vba, $inner.Vba, $verdict)

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) disagreed with Application.Caller" }
    Complete-Test -Pass -Detail "tracer reproduced Application.Caller in every case; nested Change was $(if ($inner.Vba -eq $outer.Vba) { 'inherited' } else { 'reset' })"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
