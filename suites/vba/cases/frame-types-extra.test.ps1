# The remaining declarable shapes: class types, fixed-length strings, typed and
# multi-dimension arrays, UDT arrays, nested Types, typed Optional defaults.
$case = @{ Name='frame-types-extra'
     Setup=@'
Public Type T_Inner
    Code As Long
    Tag As String * 8
End Type
Public Type T_Outer
    Head As T_Inner
    Total As Double
End Type
Public Sub T_ClsRange(ByVal r As Range)
    Dim z As Long
    z = r.Row
End Sub
Public Sub T_ClsSheet(ByVal w As Worksheet)
    Dim z As String
    z = w.Name
End Sub
Public Sub T_FixStr(ByRef p As T_Inner)
    Dim z As String
    z = p.Tag
End Sub
Public Sub T_NestUdt(ByRef p As T_Outer)
    Dim z As Double
    z = p.Total
End Sub
Public Sub T_ArrDbl(ByRef a() As Double)
    Dim z As Double
    z = a(1)
End Sub
Public Sub T_ArrVar(ByRef a() As Variant)
    Dim z As Variant
    z = a(1)
End Sub
Public Sub T_ArrUdt(ByRef a() As T_Inner)
    Dim z As Long
    z = a(1).Code
End Sub
Public Sub T_VarHoldsArr(ByVal v As Variant)
    Dim z As Long
    z = UBound(v)
End Sub
Public Sub T_OptStr(Optional ByVal s As String = "OPTDEF")
    Dim z As String
    z = s
End Sub
Public Sub T_OptDbl(Optional ByVal d As Double = 3.5)
    Dim z As Double
    z = d
End Sub
Public Sub T_XtraDrive()
    Dim inner As T_Inner
    Dim outer As T_Outer
    Dim ad(1 To 3) As Double
    Dim av(1 To 3) As Variant
    Dim au(1 To 2) As T_Inner
    Dim packed As Variant
    inner.Code = &H51515151
    inner.Tag = "TAGVAL"
    outer.Head.Code = &H52525252
    outer.Total = 987.25
    ad(1) = 1.25
    ad(2) = 2.5
    av(1) = "AV1"
    av(2) = 7
    au(1).Code = &H53535353
    packed = Array(11, 22, 33)
    Call T_ClsRange(ThisWorkbook.Worksheets(1).Range("A1"))
    Call T_ClsSheet(ThisWorkbook.Worksheets(1))
    Call T_FixStr(inner)
    Call T_NestUdt(outer)
    Call T_ArrDbl(ad)
    Call T_ArrVar(av)
    Call T_ArrUdt(au)
    Call T_VarHoldsArr(packed)
    Call T_OptStr
    Call T_OptStr("GIVEN")
    Call T_OptDbl
    Call T_OptDbl(9.75)
End Sub
'@
     Invoke=@{ Name='T_XtraDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 11) {
            return "expected 11 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        # A specific class type is still an Object at the p-code level -- the
        # class NAME is compile-time, exactly as an Enum's is.
        foreach ($fn in @('T_ClsRange','T_ClsSheet')) {
            if ($entry[$fn].typetext -ne 'Object') {
                return "$fn signature was [$($entry[$fn].typetext)], expected Object" }
        }
        # Optional with a non-Variant default: the CALLER materialises the
        # default, so it reads as an ordinary value of the declared type.
        if ($entry['T_OptStr'].args -ne 'a1:String="OPTDEF"') {
            return "T_OptStr omitted-default args were [$($entry['T_OptStr'].args)]" }
        if ($entry['T_OptDbl'].args -ne 'a1:Double=3.5') {
            return "T_OptDbl omitted-default args were [$($entry['T_OptDbl'].args)]" }
        $missing = @()
        foreach ($fn in @('T_ClsRange','T_ClsSheet','T_FixStr','T_NestUdt',
                          'T_ArrDbl','T_ArrVar','T_ArrUdt','T_VarHoldsArr',
                          'T_OptStr','T_OptDbl')) {
            if (-not $entry.ContainsKey($fn)) { $missing += "$fn=MISSING"; continue }
            if ($entry[$fn].typetext -match '\?') {
                $missing += "$fn$($entry[$fn].typetext)" }
        }
        if ($missing.Count -gt 0) {
            return "no recoverable type for: $($missing -join ' ')" }
        $null }
     Why='specific classes, fixed-length strings, typed and UDT arrays, nested
          Types and non-Variant Optional defaults -- the declarable shapes the
          first two passes left out' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
