# `caller` is what Application.Caller returns on the procedure's first line. Compared against
# the real call rather than predicted, since the documented table says nothing about nesting or
# events. Each procedure reads it inline: whether the call stack matters is part of the question.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

# the formatter mirrors src/xll/caller.cpp's vocabulary: the test is a string comparison
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

    # one arming session per trigger, so each trigger's rows stand alone in their own trace file
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

    function Get-TracerCaller($rows, [string]$proc) {
        $r = @($rows | Where-Object { $_.function -eq $proc })
        if ($r.Count -eq 0) { return $null }
        $x = $r[0]
        # translated here, so the VBA side stays an independent witness; IsArray covers a
        # toolbar and a menu too
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
    # the tracer writes [Book]Sheet; VBA's Parent.Name is the bare sheet
    function Same([string]$tracer, [string]$vba) {
        if ($null -eq $tracer -or $null -eq $vba) { return $false }
        if ($tracer -eq $vba) { return $true }
        if ($tracer -like 'cell|*' -and $vba -like 'cell|*') {
            $t = $tracer.Split('|'); $v = $vba.Split('|')
            return ($t[1] -eq $v[1]) -and ($t[2] -like "*$($v[2])*")
        }
        return $false
    }

    # a script variable, not the pipeline: assigning the result would swallow Check's output
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

    # ---- 4. nested: Run -> XRWriteCell -> Worksheet_Change -----------------
    $r4 = Invoke-Trigger {
        $app.EnableEvents = $true
        $app.Run($bang + 'XRWriteCell') | Out-Null
        $app.EnableEvents = $false
    }
    Compare-Proc $r4 'XRWriteCell'      'nested-outer-agrees'
    $outer = $script:lastCmp
    Compare-Proc $r4 'Worksheet_Change' 'nested-inner-agrees'
    $inner = $script:lastCmp

    # or the comparison above passed vacuously on a nesting that never happened
    $both = ($null -ne $outer.Tracer) -and ($null -ne $inner.Tracer)
    Check 'nested-both-frames-traced' $both "outer='$($outer.Tracer)' inner='$($inner.Tracer)'"

    # 5. Run and an event both answer #REF!, so case 4 cannot tell inherited from reset; Auto_Open
    # answers the document name. RunAutoMacros, since an automated Workbooks.Open skips Auto_Open.
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

    # recorded, not asserted: only a distinctive outer caller can tell the two apart
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
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
