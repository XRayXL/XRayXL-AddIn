# WHERE ELSE CAN AN ERROR COME FROM? -- class modules, constructors, forms.
#
# The error chain was built and proved against ordinary Subs in a standard
# module. Five other shapes reach the interpreter by a different route, and each
# is a plausible way the chain breaks:
#
#   class method       an ordinary call on an object -- the control
#   Class_Initialize   invoked by the interpreter as part of `New`, NOT by a
#                      call opcode. If it opens no frame the way a Sub does, the
#                      raise lands with nothing on the shadow stack and is held
#                      for the next frame -- which would be the CALLER, giving
#                      it somebody else's throw.
#   Property Get       a different call shape again
#   Property Let       and so is assignment
#   Class_Terminate    driven by refcount, so it can fire anywhere -- including
#                      somewhere with no relation to the code that dropped the
#                      last reference
#   UserForm method    a form module is a class with a designer attached
#
# EXPLORATORY, AND HONEST ABOUT IT. The chain for these is NOT known in advance,
# so the per-shape outcome is REPORTED rather than asserted. What IS asserted is
# the part that must hold whatever the interpreter does:
#
#   - the catcher resumed, so it must read `handled`
#   - every exit row carries an outcome
#   - a raise is either attributed to a frame or COUNTED as unattributable
#     (errNoFrame); silently vanishing is the one outcome that is a bug
#
# That way a surprise shows up as a reported difference instead of a red test
# asserting a guess.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$classCode = @'
Private mV As Long

Private Sub Class_Initialize()
    Dim a As Long
    a = 1
    If gRaiseIn = "init" Then Err.Raise 5, "XRayCase", "raised in Class_Initialize"
    a = 2
End Sub

Private Sub Class_Terminate()
    Dim b As Long
    b = 1
    If gRaiseIn = "term" Then Err.Raise 5, "XRayCase", "raised in Class_Terminate"
    b = 2
End Sub

Public Sub Method()
    Dim c As Long
    c = 1
    If gRaiseIn = "method" Then Err.Raise 5, "XRayCase", "raised in a class method"
    c = 2
End Sub

Public Property Get Value() As Long
    Dim d As Long
    d = 1
    If gRaiseIn = "get" Then Err.Raise 5, "XRayCase", "raised in Property Get"
    Value = 7
End Property

Public Property Let Value(ByVal v As Long)
    Dim e As Long
    e = 1
    If gRaiseIn = "let" Then Err.Raise 5, "XRayCase", "raised in Property Let"
    mV = v
End Property
'@

$formCode = @'
Public Sub FormMethod()
    Dim f As Long
    f = 1
    If gRaiseIn = "form" Then Err.Raise 5, "XRayCase", "raised in a form module"
    f = 2
End Sub
'@

$moduleCode = @'
Public gRaiseIn As String

' The catcher, in every case. It RESUMES, so it must read handled.
Public Sub C_Outer(ByVal which As String)
    Dim n As Long
    gRaiseIn = which
    On Error GoTo Caught
    n = 1
    C_Drive which
    n = 2
    gRaiseIn = ""
    Exit Sub
Caught:
    n = 3
    n = 4
    gRaiseIn = ""
End Sub

Public Sub C_Drive(ByVal which As String)
    Dim o As Object
    Dim x As Long
    Select Case which
        Case "init"
            Set o = New XRClass
        Case "term"
            Set o = New XRClass
            Set o = Nothing
        Case "method"
            Set o = New XRClass
            o.Method
        Case "get"
            Set o = New XRClass
            x = o.Value
        Case "let"
            Set o = New XRClass
            o.Value = 3
        Case "form"
            XRForm.FormMethod
    End Select
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # no MSForms designer means SKIP, not a pass on the other five shapes
    New-XRayMacroBook $sx 'ErrShapes' @(
        @{ Kind=1; Name='ErrShapes'; Code=$moduleCode }
        @{ Kind=2; Name='XRClass'; Code=$classCode }
        @{ Kind=3; Name='XRForm';    Code=$formCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $shapes = @('method','init','get','let','term','form')

    $dlgBefore = @(Get-SessionDialogs).Count
    $report = @()
    foreach ($which in $shapes) {
        $mark = Get-LogLength $paths.Log
        [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
        $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
        if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }
        # an unverified raise slot is a product failure, not a SKIP: every shape would read `returned`
        if ($armLine -match 'NO ERROR ATTRIBUTION') {
            Complete-Test -Fail -Detail ("the raise slot did not verify on this VBE7, so the " +
                                         "'$which' shape cannot be attributed: $armLine")
        }

        $threw = ''
        try { $app.Run($leaf + '!C_Outer', $which) | Out-Null }
        catch { $threw = $_.Exception.Message }
        $lossy = Stop-XRayTrace $sx
        if ($lossy) { Complete-Test -Fail -Detail $lossy }

        $totals = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark 20
        $rows   = @(Read-TraceRows $sx.ProcId)
        $exits  = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })

        # the chain, innermost first -- exits are written as frames close
        $chain = @($exits | ForEach-Object {
            "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })"
        })
        $nThrew   = @($exits | Where-Object { $_.outcome -eq 'threw' }).Count
        $nHandled = @($exits | Where-Object { $_.outcome -eq 'handled' }).Count
        $noOutcome = @($exits | Where-Object { -not $_.outcome }).Count
        $errNoFrame = 0
        if ($totals -match 'errNoFrame=(\d+)') { $errNoFrame = [int]$Matches[1] }
        $raises = 0
        if ($totals -match ' raises=(\d+)')   { $raises = [int]$Matches[1] }

        $report += [pscustomobject]@{
            Shape = $which; Chain = ($chain -join ' -> ')
            Threw = $nThrew; Handled = $nHandled; NoOutcome = $noOutcome
            Raises = $raises; NoFrame = $errNoFrame; Ex = $threw
        }

        # ---- what must hold whatever the interpreter does ------------------
        #
        # EXCEPT FOR Class_Terminate, WHICH VBA DOES NOT PROPAGATE. Measured:
        # raising there produced three message boxes (the harness dismissed
        # them) and no handler ever ran, so there is nothing to mark `handled`.
        # That is Excel's behaviour, not the tracer's -- the row still reads
        # `threw`, which is the honest answer, and asserting a catcher here
        # would be asserting something VBA does not do.
        if ($which -ne 'term') {
            Check "$which-catcher-handled" ($nHandled -ge 1) `
                  "chain: $($chain -join ' -> ')"
        } else {
            Check "$which-not-propagated-by-vba" ($nHandled -eq 0) `
                  "a raise out of Class_Terminate never reaches a handler; handled=$nHandled, chain: $($chain -join ' -> ')"
        }
        Check "$which-every-exit-has-an-outcome" ($noOutcome -eq 0) `
              "rows without outcome: $noOutcome of $($exits.Count)"
        # A raise is attributed to SOME frame, or counted as unattributable.
        # Vanishing silently is the only bug here.
        Check "$which-raise-accounted-for" (($nThrew -ge 1) -or ($errNoFrame -ge 1)) `
              "threw=$nThrew errNoFrame=$errNoFrame raises=$raises  chain: $($chain -join ' -> ')"
    }

    Write-Output ''
    Write-Output 'PER SHAPE -- reported, not asserted:'
    foreach ($r in $report) {
        Write-Output ("  {0,-7} raises={1} threw={2} handled={3} noFrame={4}  {5}" -f `
            $r.Shape, $r.Raises, $r.Threw, $r.Handled, $r.NoFrame, $r.Chain)
        if ($r.Ex) { Write-Output ("            COM said: {0}" -f ($r.Ex -replace '\s+',' ')) }
    }


    # Raising inside Class_Terminate makes VBA put up message boxes rather than
    # propagate, so the dialogs are expected here and asserted, not waived.
    $dlg = @(Get-SessionDialogs)
    $dlgNew = $dlg.Count - $dlgBefore
    if ($dlgNew -gt 0) { Write-DialogsHandled }
    Check 'class-terminate-raised-message-boxes' ($dlgNew -ge 1) `
          ("{0} dialog(s) during the shapes{1}" -f $dlgNew, $(if ($dlgNew -gt 0) { ', last: ' + (($dlg | Select-Object -Last 1) -join '') } else { '' }))

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed across $($shapes.Count) shapes" }
    Complete-Test -Pass -Detail (($report | ForEach-Object { "$($_.Shape):threw=$($_.Threw)/handled=$($_.Handled)" }) -join ' ')
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
