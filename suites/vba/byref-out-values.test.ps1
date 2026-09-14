# ByRef ARGUMENTS A PROCEDURE CHANGED -- does the exit row show the new value?
#
# ByRef is VBA's DEFAULT. `Sub Calc(result As Double)` that fills in `result`
# is ordinary code, and until now its entire effect was invisible: the entry row
# captured the arguments at the FIRST STATEMENT, deliberately, so the "before"
# value is what a reader saw, and there was no "after" anywhere.
#
# THE RISK THIS MEASURES FIRST. Argument types come from the LOAD opcode -- a
# parameter the body never READS emits no typed load and so has no recoverable
# type. A pure out-parameter is exactly that shape:
#
#     Sub FillOnly(result As Double)      ' writes it, never reads it
#         result = 9.75
#     End Sub
#
# If a write emits no usable type, the decoder cannot know the parameter is
# ByRef, and the most common out-parameter in VBA would be the one case the
# feature missed. So the two shapes are separated and both are reported:
#
#     FillOnly     assigns only          -- type recoverable?
#     ReadWrite    reads, then assigns   -- type certainly recoverable
#
# AND THE NEGATIVE CONTROL MATTERS AS MUCH AS THE POSITIVE. A ByVal parameter
# whose copy the callee modified must NOT be reported: the caller never sees
# that change, so an "after" value would assert an effect that does not exist.
. (Join-Path $PSScriptRoot '..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$moduleCode = @'
Public gSeen As Double

' Read back AFTER disarm, so it is VBA's own account of what the caller saw
' rather than anything the tracer produced.
Public Function GetSeen() As Double
    GetSeen = gSeen
End Function

' Writes its parameter and never reads it -- the pure out-parameter.
Public Sub FillOnly(result As Double)
    result = 9.75
End Sub

' Reads it and then writes it, so a typed load certainly exists.
Public Sub ReadWrite(v As Double)
    Dim t As Double
    t = v
    v = t + 100
End Sub

' ByRef, and deliberately does NOT touch it.
Public Sub NoTouch(v As Double)
    Dim t As Double
    t = v
End Sub

' ByVal. The copy changes; the CALLER never sees it. Must not be reported.
Public Sub ByValChanged(ByVal v As Double)
    Dim t As Double
    t = v
    v = t + 100
End Sub

' A ByRef string, to prove it is not a numeric-only path.
Public Sub FillString(s As String)
    Dim t As String
    t = s
    s = t & "-out"
End Sub

' WRITE-ONLY STRING -- the one shape still not recovered. Assigned and never
' read, so no typed load names the slot; unlike the numeric types, no typed
' STORE reaches it either (measured in isolation).
Public Sub FillStringOnly(s As String)
    s = "written"
End Sub

' ByVal beside ByRef: the ByVal copy changes, the ByRef argument does not.
Public Sub MixedByValOnly(ByVal n As Long, r As Long)
    Dim t As Long
    t = r
    n = n + 1
End Sub

' Both change; only the ByRef one reaches the caller.
Public Sub MixedBoth(ByVal n As Long, r As Long)
    n = n + 1
    r = r + 10
End Sub

Public Sub Drive()
    Dim a As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim s As String

    a = 1.5:  FillOnly a
    b = 2.5:  ReadWrite b
    c = 3.5:  NoTouch c
    d = 4.5:  ByValChanged d
    s = "in": FillString s
    Dim s2 As String
    s2 = "in2": FillStringOnly s2
    Dim m1 As Long
    Dim m2 As Long
    m1 = 5: MixedByValOnly 3, m1
    m2 = 5: MixedBoth 3, m2

    ' Proof from VBA's own side that the caller really did see the change --
    ' otherwise the trace could be right about nothing.
    gSeen = a
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ByRef' @(
        @{ Kind=1; Name='ByRefCase'; Code=$moduleCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'   'TRUE')
    [void](Set-XRayTraceParam $sx 'VBA' 'RETVAL' 'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH'  'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $totals = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark 20
    $rows   = @(Read-TraceRows $sx.ProcId)

    Write-Output ''
    Write-Output 'before and after, per procedure:'
    Write-Output ('  {0,-14} {1,-34} {2}' -f 'procedure', 'ENTRY (before)', 'EXIT (after, only if changed)')
    foreach ($f in @('FillOnly','ReadWrite','NoTouch','ByValChanged','FillString','FillStringOnly','MixedByValOnly','MixedBoth')) {
        Write-Output ('  {0,-14} {1,-34} {2}' -f $f, (ArgsOf $rows $f),
                      $(if ((ExitArgsOf $rows $f)) { ExitArgsOf $rows $f } else { '(none)' }))
    }
    Write-Output ''
    foreach ($k in @('byrefEligible','byrefChanged','byrefSame','byrefDeclined')) {
        if ($totals -match "$k=(\d+)") { Write-Output ("  {0,-14} {1}" -f $k, $Matches[1]) }
    }

    # ---- VBA'S OWN VIEW, first: the change really happened -----------------
    # Without this the trace could be consistent and still describing nothing.
    $seen = [double]$app.Run($leaf + '!GetSeen')
    Check 'the-caller-really-saw-the-change' ($seen -eq 9.75) `
          "VBA's gSeen after FillOnly = $seen (expected 9.75)"

    # ---- THE READ-THEN-WRITE SHAPE, where the type is certainly known ------
    Check 'a-changed-byref-argument-is-reported-at-the-exit' `
          ((ExitArgsOf $rows 'ReadWrite') -match '102\.5') `
          ("ReadWrite entry='" + (ArgsOf $rows 'ReadWrite') + "' exit='" + (ExitArgsOf $rows 'ReadWrite') + "' (2.5 -> 102.5)")

    Check 'the-entry-row-still-shows-the-value-going-in' `
          ((ArgsOf $rows 'ReadWrite') -match '2\.5') `
          ("ReadWrite entry='" + (ArgsOf $rows 'ReadWrite') + "'")

    Check 'a-changed-byref-string-is-reported-too' `
          ((ExitArgsOf $rows 'FillString') -match 'in-out') `
          ("FillString entry='" + (ArgsOf $rows 'FillString') + "' exit='" + (ExitArgsOf $rows 'FillString') + "'")

    # ---- UNCHANGED MEANS SILENT -------------------------------------------
    # An exit row with args means "these moved". A procedure that touched
    # nothing must not produce one, or the column stops meaning anything.
    Check 'an-untouched-byref-argument-is-not-reported' `
          (-not (ExitArgsOf $rows 'NoTouch')) `
          ("NoTouch exit args='" + (ExitArgsOf $rows 'NoTouch') + "' (expected empty)")

    # ---- THE NEGATIVE CONTROL ---------------------------------------------
    # ByValChanged modifies its COPY. The caller never sees it. Reporting it
    # would assert an effect that does not exist.
    Check 'a-byval-copy-the-callee-changed-is-not-reported' `
          (-not (ExitArgsOf $rows 'ByValChanged')) `
          ("ByValChanged exit args='" + (ExitArgsOf $rows 'ByValChanged') + "' (the caller never sees this)")

    # ---- ByVal BESIDE ByRef -------------------------------------------------
    # Any ByRef parameter makes the exit re-read happen, so what moved must be
    # judged per parameter or the ByVal copy is reported with it.
    Check 'a-byval-copy-beside-an-untouched-byref-is-not-reported' `
          (-not (ExitArgsOf $rows 'MixedByValOnly')) `
          ("MixedByValOnly entry='" + (ArgsOf $rows 'MixedByValOnly') + "' exit='" + (ExitArgsOf $rows 'MixedByValOnly') + "' (expected empty)")

    $mixed = ExitArgsOf $rows 'MixedBoth'
    Check 'only-the-byref-argument-is-reported-when-both-change' `
          (($mixed -match '(^| )a2:[^ ]*=15$') -and ($mixed -notmatch 'a1:')) `
          ("MixedBoth entry='" + (ArgsOf $rows 'MixedBoth') + "' exit='$mixed' (expected only a2, reading 15)")

    # ---- THE CONTRACT: exit rows carry VALUES, not the signature -----------
    $exitRows = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })
    $leaky = @($exitRows | Where-Object { $_.argcount -or $_.typetext })
    Check 'exit-rows-do-not-repeat-argcount-or-typetext' ($leaky.Count -eq 0) `
          ("rows carrying them: " + (@($leaky | ForEach-Object { "$($_.function) argcount='$($_.argcount)' typetext='$($_.typetext)'" }) -join ' | '))

    # ---- THE COUNTERS AGREE WITH THE ROWS ---------------------------------
    $tChanged = 0
    if ($totals -match 'byrefChanged=(\d+)') { $tChanged = [int]$Matches[1] }
    $rChanged = @($exitRows | Where-Object { $_.args }).Count
    Check 'the-changed-counter-agrees-with-the-rows' ($tChanged -eq $rChanged) `
          "byrefChanged=$tChanged, exit rows carrying args=$rChanged"

    # ---- THE PURE OUT-PARAMETER, which is the commonest shape of all -------
    #
    # `FillOnly` writes its parameter and never reads it, so there is no typed
    # LOAD: the type comes from the store (`op - 32`). Both directions are
    # checked -- the value written, and what went in.
    Check 'a-pure-out-parameter-reports-its-value' `
          ((ExitArgsOf $rows 'FillOnly') -match '9\.75') `
          ("FillOnly entry='" + (ArgsOf $rows 'FillOnly') + "' exit='" + (ExitArgsOf $rows 'FillOnly') + "' (writes 9.75, never reads it)")

    Check 'a-pure-out-parameter-still-shows-what-went-in' `
          ((ArgsOf $rows 'FillOnly') -match '1\.5') `
          ("FillOnly entry='" + (ArgsOf $rows 'FillOnly') + "' (caller passed 1.5)")

    # ---- THE WRITE-ONLY STRING, ASSIGNED A LITERAL ---------------------------
    # Its typed load sits past an opcode with no derivable length, so the type
    # may stay unrecovered; a wrong type is never acceptable. Declared ByRef
    # String, the two honest readings are String& and ?.
    $sigOnly = SigOf $rows 'FillStringOnly'
    Check 'a-write-only-string-is-typed-right-or-not-at-all' ($sigOnly -in @('String&', '?')) `
          "FillStringOnly sig='$sigOnly', expected String& or ?"

    Check 'a-write-only-string-reports-what-it-wrote' `
          ((ExitArgsOf $rows 'FillStringOnly') -match 'written') `
          ("FillStringOnly entry='" + (ArgsOf $rows 'FillStringOnly') +
           "' exit='" + (ExitArgsOf $rows 'FillStringOnly') + "'")


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail "changed=$tChanged reported at the exit; ByVal and untouched stayed silent"
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
