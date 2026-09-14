# ORDINARY TRACING OUT OF CLASS MODULES AND FORMS -- no errors involved.
#
# Every VBA case in the suite drives Subs and Functions in a STANDARD module.
# That is one container out of three, and the other two reach the interpreter
# differently: a class module is instantiated and its procedures called through
# an object, a form module is a class with a designer attached, and constructors
# and destructors are invoked by the interpreter rather than by a call opcode.
#
# The error work needed those shapes and so exercised them by accident. This
# asserts the ORDINARY behaviour on purpose, because "the outcome column is
# right" is a much weaker claim than "the rows are there at all, named, nested
# and paired".
#
# WHAT IS ASSERTED, and why each is a way it could be wrong rather than a box
# ticked:
#
#   present   a procedure that ran but produced no row is invisible, which is
#             the whole failure mode this tool exists to prevent
#   named     `Method` and `Value` must be NAMED, not reported as a trailer
#             address -- identity resolution walks different structures for a
#             class than for a standard module
#   paired    entry and exit by SPAN, in both directions: an unpaired entry
#             reads as a hang, an unpaired exit as a call that never happened
#   nested    Class_Initialize runs INSIDE the caller that said `New`, so its
#             parent must be that frame and not 0
#   returns   the typed Functions and the Property GET carry a value; the
#             Subs and the Property LET carry none -- both halves asserted
#   outcome   all of it `returned` -- the negative control for the error work
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$classCode = @'
Private mV As Long

Private Sub Class_Initialize()
    mV = 10
End Sub

Private Sub Class_Terminate()
    mV = 0
End Sub

Public Sub Method()
    Dim c As Long
    c = mV + 1
End Sub

Public Function Doubled(ByVal n As Long) As Long
    Doubled = n * 2
End Function

Public Property Get Value() As Long
    Value = mV
End Property

Public Property Let Value(ByVal v As Long)
    mV = v
End Property
'@

$formCode = @'
Public Function FormAdd(ByVal a As Long, ByVal b As Long) As Long
    FormAdd = a + b
End Function
'@

$moduleCode = @'
Public gLast As Long

Public Sub N_Drive()
    Dim o As New XRNClass
    Dim x As Long
    o.Method
    x = o.Doubled(21)
    o.Value = 5
    x = o.Value
    gLast = XRNForm.FormAdd(2, 3)
    Set o = Nothing
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # no MSForms designer means SKIP, not a half-run pass
    New-XRayMacroBook $sx 'ClsForm' @(
        @{ Kind=1; Name='NormCase'; Code=$moduleCode }
        @{ Kind=2; Name='XRNClass'; Code=$classCode }
        @{ Kind=3; Name='XRNForm';  Code=$formCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!N_Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows    = @(Read-TraceRows $sx.ProcId)
    $entries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') })
    $exits   = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    $names   = @($entries | ForEach-Object { $_.function })

    # ---- present ----------------------------------------------------------
    $wantClass = @('Class_Initialize','Method','Doubled','Value','Class_Terminate')
    $missing = @($wantClass | Where-Object { $names -notcontains $_ })
    Check 'class-procedures-traced' ($missing.Count -eq 0) `
          ("missing: " + ($missing -join ',') + "   saw: " + ($names -join ','))

    Check 'form-procedure-traced' ($names -contains 'FormAdd') ("saw: " + ($names -join ','))

    # ---- named, not an address --------------------------------------------
    # Identity resolution walks a different structure for a class than for a
    # standard module; a procedure it cannot name is reported as its trailer
    # address, which is honest but useless.
    $unnamed = @($entries | Where-Object { $_.function -match '^0x[0-9A-F]+$' })
    Check 'no-procedure-reported-as-an-address' ($unnamed.Count -eq 0) `
          ("unnamed: " + (@($unnamed | ForEach-Object { $_.function }) -join ','))

    # ---- paired, both directions ------------------------------------------
    $eSpans = @($entries | ForEach-Object { $_.span })
    $xSpans = @($exits   | ForEach-Object { $_.span })
    $orphanExits   = @($exits   | Where-Object { $eSpans -notcontains $_.span })
    $orphanEntries = @($entries | Where-Object { $xSpans -notcontains $_.span })
    Check 'no-exit-without-its-entry' ($orphanExits.Count -eq 0) `
          ("orphans: " + (@($orphanExits | ForEach-Object { $_.function }) -join ','))
    Check 'no-entry-without-its-exit' ($orphanEntries.Count -eq 0) `
          ("orphans: " + (@($orphanEntries | ForEach-Object { $_.function }) -join ','))

    # ---- nested under the caller that constructed it ----------------------
    # Class_Initialize is invoked by the interpreter as part of New, not by a
    # call opcode. If that produced a top-level frame its parent would be 0.
    $init = @($entries | Where-Object { $_.function -eq 'Class_Initialize' })
    $initParent = if ($init.Count) { [string]$init[0].parent } else { '' }
    Check 'constructor-nests-under-its-caller' (($init.Count -ge 1) -and ($initParent -ne '0')) `
          "Class_Initialize parent: $(if ($init.Count) { $init[0].parent } else { '(absent)' })"

    # ---- RETURN VALUES OUT OF A CLASS AND A FORM --------------------------
    #
    # Class and form exits carry no type, so the result is typed from the store
    # opcode before them. Both halves: a Sub and a Property Let have no result,
    # and their local at the same frame offset must not be reported as one.
    #
    # The two `Value` rows share a name: Property Let ran first (o.Value = 5),
    # Property Get second (x = o.Value). They are told apart by span order, not
    # by name, because the name is all VBA gives them.
    Check 'class-function-returns-its-value' ((RetOf $rows 'Doubled') -eq '42') `
          ("Doubled ret='" + (RetOf $rows 'Doubled') + "' (Doubled(21) = 42)")

    Check 'form-function-returns-its-value' ((RetOf $rows 'FormAdd') -eq '5') `
          ("FormAdd ret='" + (RetOf $rows 'FormAdd') + "' (FormAdd(2,3) = 5)")

    # A Property GET is a Function in every way that matters here.
    $valRows = @($exits | Where-Object { $_.function -eq 'Value' } | Sort-Object { [int]$_.span })
    $valGet  = if ($valRows.Count -ge 2) { [string]$valRows[1].ret } else { '(missing)' }
    $valLet  = if ($valRows.Count -ge 2) { [string]$valRows[0].ret } else { '(missing)' }
    Check 'property-get-returns-its-value' ($valGet -eq '5') `
          ("Property Get ret='$valGet' (mV was set to 5)")
    Check 'property-let-reports-no-return-value' ($valLet -eq '') `
          ("Property Let ret='$valLet' -- a Let has no result to report")

    # ...and neither does a Sub. A local sitting at the result offset must not
    # be mistaken for one.
    $subNames = @('N_Drive','Method','Class_Initialize','Class_Terminate')
    $subsWithRet = @($exits | Where-Object { ($subNames -contains $_.function) -and $_.ret })
    Check 'subs-report-no-return-value' ($subsWithRet.Count -eq 0) `
          ("subs carrying a value: " + (@($subsWithRet | ForEach-Object { "$($_.function)='$($_.ret)'" }) -join ','))

    # ---- the negative control for the error work --------------------------
    $notReturned = @($exits | Where-Object { $_.outcome -ne 'returned' })
    Check 'every-outcome-is-returned' ($notReturned.Count -eq 0) `
          ("not returned: " + (@($notReturned | ForEach-Object {
              "$($_.function)=$(if ($_.outcome) { $_.outcome } else { '(none)' })" }) -join ','))

    Write-Output ''
    Write-Output 'what was traced:'
    foreach ($e in $entries) { Write-Output ("  entry span={0,-3} {1,-18} depth={2} parent={3}" -f $e.span, $e.function, $e.depth, $e.parent) }
    foreach ($x in $exits)   { Write-Output ("  exit  span={0,-3} {1,-18} ret='{2}' ticks={3} trust={4}" -f $x.span, $x.function, $x.ret, $x.ticks, $x.trust) }


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "$($entries.Count) entries / $($exits.Count) exits across class, form and standard modules"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
