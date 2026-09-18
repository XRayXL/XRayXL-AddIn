# The p-code length table fits this build. A wrong length does not throw: it steps into the
# middle of an operand and types the wrong parameter with confidence. Two facts:
#
#    1. The opcode set is the one the table describes. The dispatch table's equivalence
#       partition (for every slot, the lowest slot sharing its handler) names no address,
#       so it survives rebasing, and the arm log says whether it matched.
#    2. The table walks real compiler output cleanly. The case runs varied VBA (typed and
#       Variant arithmetic, arrays, objects, a Property, error handling, loops, string
#       work) and requires every procedure to walk from offset 0 to a clean exit with no
#       resynchronisation.
#
# AsLoaded, from a saved workbook: p-code compiled by Excel on open is what a user has.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$src = @'
Public Event Ping()
Private mTotal As Long

Public Property Get Total() As Long
    Total = mTotal
End Property
Public Property Let Total(ByVal v As Long)
    mTotal = v
End Property

Public Function Typed(ByVal a As Long, ByVal b As Double, ByVal s As String) As Double
    Dim i As Integer, c As Currency, by As Byte, sg As Single, bo As Boolean
    i = a: c = b: by = 1: sg = b: bo = (a > 0)
    Typed = a + b + Len(s) + i + c + by + sg
End Function

Public Function Varied(ByVal v As Variant) As Variant
    Dim w As Variant, n As Long
    w = v
    w = w + 1: w = w - 1: w = w * 2: w = w / 2: w = -w
    If w > 0 And w < 1000 Then n = 1
    If w = 0 Or w <> 1 Then n = n + 1
    Varied = w & CStr(n)
End Function

Public Function Arrays(ByVal n As Long) As Long
    Dim a() As Long, v As Variant, i As Long, t As Long
    ReDim a(1 To n)
    For i = 1 To n
        a(i) = i * i
    Next i
    ReDim Preserve a(1 To n + 1)
    v = a
    For Each v In a
        t = t + 1
    Next
    Erase a
    Arrays = t
End Function

Public Function Objects() As Long
    Dim c As New Collection, o As Object, n As Long
    c.Add 1: c.Add 2
    Set o = c
    n = c.Count + o.Count
    With c
        n = n + .Count
    End With
    Objects = n
End Function

Public Function Strings(ByVal s As String) As String
    Dim t As String, i As Long
    t = ""
    For i = 1 To 3
        t = t & Mid(s, i, 1) & "-"
    Next i
    If t Like "*-*" Then t = UCase(t)
    Strings = Trim(t)
End Function

Public Function Guarded(ByVal n As Long) As Long
    On Error GoTo Bad
    If n = 0 Then Err.Raise 5
    Guarded = 100 \ n
    Exit Function
Bad:
    Guarded = -1
End Function

Public Sub Drive()
    Dim d As Double, v As Variant, s As String
    Total = 7
    d = Typed(2, 1.5, "abc")
    v = Varied(3)
    d = d + Arrays(4) + Objects() + Guarded(2) + Guarded(0) + Total
    s = Strings("hello")
    RaiseEvent Ping
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $driveCode = @'
Public Sub Drive()
    Dim c As New CFit
    c.Drive
End Sub
'@
    # A CLASS, not a standard module: Property, Event and RaiseEvent only
    # compile in one, and they are three of the shapes being asserted.
    New-XRayMacroBook $sx 'pcode' @(
        @{ Kind = 2; Name = 'CFit'; Code = $src }
        @{ Kind = 1; Name = 'MFit'; Code = $driveCode }
    ) -Leaf 'PcodeFit.xlsm'

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if (-not $armLine) { Complete-Test -Fail -Detail 'VBA never reported an arm outcome' }

    # ---- 1. the opcode set ------------------------------------------------
    # The derivation line carries the partition either way, so the assertion
    # can tell "matched" from "the line never appeared".
    $partLine = @(Get-Content $paths.Log | Select-Object -Skip $mark |
                  Select-String 'partition=0x') | Select-Object -Last 1
    Check 'partition-line-is-reported' ($null -ne $partLine) 'no partition= in the arm log'
    if ($partLine) {
        $txt = [string]$partLine.Line
        Check 'opcode-set-is-the-one-the-table-was-measured-against' `
              ($txt -match 'known opcode set') `
              (($txt -split '\s+' | Where-Object { $_ -like 'partition=*' }) -join ' ')
    }
    # A degraded session says so in as many words; types would be absent.
    $unknownSet = @(Get-Content $paths.Log | Select-Object -Skip $mark |
                    Select-String 'UNKNOWN OPCODE SET')
    Check 'types-were-not-degraded-off' ($unknownSet.Count -eq 0) `
          (($unknownSet | ForEach-Object { $_.Line }) -join ' | ')

    # ---- run the VBA ------------------------------------------------------
    [void]$app.Run('MFit.Drive')
    $mark2 = Get-LogLength $paths.Log
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    [void](Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2)

    # ---- 2. every procedure walked cleanly --------------------------------
    $walkLine = @(Get-Content $paths.Log | Select-Object -Skip $mark2 |
                  Select-String 'procedure\(s\) walked cleanly') | Select-Object -Last 1
    Check 'walk-health-is-reported' ($null -ne $walkLine) 'no clean-walk line at disarm'
    if ($walkLine) {
        $m = [regex]::Match([string]$walkLine.Line, '(\d+) of (\d+) procedure\(s\) walked cleanly')
        $clean = [int]$m.Groups[1].Value; $walks = [int]$m.Groups[2].Value
        # Something must have been walked, or "100%" is vacuous.
        Check 'the-vba-was-actually-walked' ($walks -ge 8) "walks=$walks"
        Check 'every-procedure-walked-cleanly' ($clean -eq $walks -and $walks -gt 0) `
              "$clean of $walks walked cleanly (offset 0 to a clean exit, no resynchronisation)"
    }

    # A length that was USED and then broke the walk is the worse defect, and
    # it names itself. Silence is the assertable state.
    $proven = @(Get-Content $paths.Log | Select-Object -Skip $mark2 |
                Select-String 'LENGTHS PROVEN WRONG')
    Check 'no-length-was-proven-wrong' ($proven.Count -eq 0) `
          (($proven | ForEach-Object { $_.Line }) -join ' | ')

    # The collapse alarm must not have fired -- and if it ever does, its own
    # words are the most useful thing this test can report.
    $collapse = @(Get-Content $paths.Log | Select-Object -Skip $mark2 |
                  Select-String 'THE LENGTH TABLE DOES NOT FIT THIS BUILD')
    Check 'table-fits-this-build' ($collapse.Count -eq 0) `
          (($collapse | ForEach-Object { $_.Line }) -join ' | ')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'known opcode set; every procedure walked cleanly'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
