# The rarer opcodes walk cleanly on this build. pcode-table-fits-this-build runs everyday VBA;
# this runs the constructs that reach opcodes ordinary code seldom does, each confirmed at its
# length by real compiled code: a variable held As IUnknown stored into a local, a record field,
# an array element, another module's Variant, a ByRef Variant and a ParamArray element -- not a
# class's own member, where VBA itself loses a reference and Excel can crash -- an Excel property assigned with named arguments, ParamArray element stores, Put
# of a fixed-length string, the file Lock and Unlock statements, Static procedures, Implements,
# Def* typing and Option Base 1 -- and a ByRef record parameter first touched through an object or
# Variant member, which the type table once could not name.
#
# Three facts: every procedure walks from offset 0 to a clean exit; no p-code warning; and the
# disarm report names no length that real code has never confirmed, so each of these opcodes is
# one the shipped table has seen walked.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$rare = @'
Option Explicit

Private Type RU
    v As Variant
End Type
Private Type RObj
    o As Object
    v As Variant
End Type
Public gU As Variant

Private Function MakeUnk() As IUnknown
    Set MakeUnk = New Collection
End Function

Private Sub SetRef(ByRef v As Variant, ByVal u As IUnknown)
    Set v = u
End Sub

Private Sub Params(ParamArray p() As Variant)
    Dim u As IUnknown
    Set u = New Collection
    p(0) = "x"
    Set p(1) = u
End Sub

Private Static Function Counter() As Long
    Dim n As Long
    n = n + 1
    Counter = n
End Function

Public Function UnkStores() As Long
    Dim u As IUnknown, v As Variant, a(1 To 2) As Variant, r As RU
    Set u = New Collection
    Set v = u
    Set v = MakeUnk()
    Set a(1) = u
    Set r.v = u
    Set r.v = MakeUnk()
    Set MRareOther.gW = u
    Set MRareOther.gW = MakeUnk()
    SetRef v, u
    Params 1, 2
    UnkStores = 1
End Function

Private Function RecObj(r As RObj) As Long
    RecObj = r.o.Count
End Function
Private Sub RecVarSet(r As RObj, ByVal x As Variant)
    r.v = x
End Sub
Public Function Records() As Long
    Dim r As RObj
    Set r.o = New Collection
    RecVarSet r, 9
    Records = RecObj(r)
End Function

Public Function NamedProperty() As Variant
    Dim rg As Range
    Set rg = ThisWorkbook.Worksheets(1).Range("A1")
    rg.Value(RangeValueDataType:=10) = 5
    ThisWorkbook.Worksheets(1).Cells(RowIndex:=1, ColumnIndex:=2) = 6
    NamedProperty = rg.Value
End Function

Public Function FileStatements() As Long
    Dim fn As Integer, p As String, fs As String * 8
    p = Environ("TEMP") & "\xrx_rare_opcodes_" & Format(Now, "hhnnss") & "_" & CLng(Rnd * 1000000) & ".dat"
    fn = FreeFile
    Open p For Random Shared As #fn Len = 16
    Lock #fn
    Unlock #fn
    Lock #fn, 1
    Unlock #fn, 1
    Lock #fn, 1 To 2
    Unlock #fn, 1 To 2
    Close #fn
    fs = "abcdefgh"
    Open p For Binary As #fn
    Put #fn, , fs
    Close #fn
    Kill p
    FileStatements = Counter() + Counter()
End Function

Public Function Shapes() As Double
    Dim s As New CRareSquare, sh As IRareShape
    s.Side = 3
    Set sh = s
    Shapes = sh.Area() + MRareOptions.OptBase(2, "a")
End Function

Public Sub Drive()
    Dim d As Double
    d = UnkStores() + NamedProperty() + FileStatements() + Shapes() + Records()
End Sub
'@

$other = @'
Option Explicit
Public gW As Variant
'@

$shape = @'
Option Explicit
Public Function Area() As Double
End Function
'@

$square = @'
Option Explicit
Implements IRareShape
Public Side As Double
Private Function IRareShape_Area() As Double
    IRareShape_Area = Side * Side
End Function
'@

$options = @'
Option Explicit
Option Base 1
DefLng L
DefStr S
DefDbl D
Public Function OptBase(ByVal lCount, ByVal sText) As Double
    Dim a(3) As Long, dSum
    a(1) = lCount: a(3) = Len(sText)
    dSum = a(1) + a(3) + LBound(a)
    OptBase = dSum
End Function
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'pcoderare' @(
        @{ Kind = 2; Name = 'IRareShape';   Code = $shape }
        @{ Kind = 2; Name = 'CRareSquare';  Code = $square }
        @{ Kind = 1; Name = 'MRareOther';   Code = $other }
        @{ Kind = 1; Name = 'MRareOptions'; Code = $options }
        @{ Kind = 1; Name = 'MRare';        Code = $rare }
    ) -Leaf 'PcodeRare.xlsm'

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    [void]$app.Run('MRare.Drive')
    $mark2 = Get-LogLength $paths.Log
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    [void](Wait-LogLine $paths.Log 'VBA trace: statements=' $mark2)
    $after = @(Get-Content $paths.Log | Select-Object -Skip $mark2)

    # ---- every procedure walked cleanly --------------------------------------
    $walkLine = @($after | Select-String 'procedure\(s\) walked cleanly') | Select-Object -Last 1
    Check 'walk-health-is-reported' ($null -ne $walkLine) 'no clean-walk line at disarm'
    if ($walkLine) {
        $m = [regex]::Match([string]$walkLine.Line, '(\d+) of (\d+) procedure\(s\) walked cleanly')
        $clean = [int]$m.Groups[1].Value; $walks = [int]$m.Groups[2].Value
        # Drive, the four callers, their helpers, the class members and the other modules.
        Check 'the-vba-was-actually-walked' ($walks -ge 12) "walks=$walks"
        Check 'every-procedure-walked-cleanly' ($clean -eq $walks -and $walks -gt 0) `
              "$clean of $walks walked cleanly (offset 0 to a clean exit, no resynchronisation)"
    }
    $pcodeWarnings = @($after | Select-String ' WARNING - VBA p-code:')
    Check 'no-p-code-warning' ($pcodeWarnings.Count -eq 0) (($pcodeWarnings | ForEach-Object { $_.Line }) -join ' | ')
    # A parameter written by an opcode the type table does not know is typed `?opNNN`.
    $untyped = @($after | Select-String ' WARNING - VBA args:')
    Check 'every-parameter-typed' ($untyped.Count -eq 0) (($untyped | ForEach-Object { $_.Line }) -join ' | ')
    $proven = @($after | Select-String 'LENGTHS PROVEN WRONG')
    Check 'no-length-was-proven-wrong' ($proven.Count -eq 0) (($proven | ForEach-Object { $_.Line }) -join ' | ')

    # ---- and every length it used is one real code has confirmed ---------------
    # The disarm report names any opcode walked this session that the shipped
    # walked table has never seen. These constructs are why those bits are set.
    $unconfirmed = @($after | Select-String 'NO RUNNING CODE HAS EVER CONFIRMED')
    Check 'every-length-used-is-confirmed' ($unconfirmed.Count -eq 0) `
          (($unconfirmed | ForEach-Object { $_.Line }) -join ' | ')

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail 'the rarer opcodes walk cleanly, each at a length real code has confirmed'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
