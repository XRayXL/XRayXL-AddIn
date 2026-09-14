# UDTs, Enums, class references, Optional and ParamArray. Weak on types,
# strict on safety: no fault, no desync, balanced frames.
$case = @{ Name='frame-udt'
     Setup=@'
Public Enum T_Colour
    ColRed = 10
    ColGreen = 20
    ColBlue = 30
End Enum
Public Type T_Point
    X As Long
    Y As Long
    Label As String
End Type
' A String first: the record's first eight bytes are a valid BSTR pointer.
Public Type T_Named
    Label As String
    N As Long
End Type
Public Sub T_UdtByRef(ByRef p As T_Point)
    Dim z As Long
    z = p.X
End Sub
Public Sub T_UdtStrFirst(ByRef p As T_Named)
    Dim z As Long
    z = p.N
End Sub
Public Sub T_EnumArg(ByVal c As T_Colour)
    Dim z As Long
    z = c
End Sub
Public Sub T_OptSupplied(Optional ByVal a As Long = 77)
    Dim z As Long
    z = a
End Sub
Public Sub T_OptVar(Optional v As Variant)
    Dim z As Long
    z = 1
End Sub
Public Sub T_ParamArr(ParamArray items() As Variant)
    Dim z As Long
    z = UBound(items)
End Sub
Public Sub T_CollArg(ByVal c As Collection)
    Dim z As Long
    z = c.Count
End Sub
Public Sub T_UdtDrive()
    Dim p As T_Point
    Dim c As Collection
    Dim q As T_Named
    q.Label = "FIRSTMEMBER"
    q.N = 7
    p.X = &H31313131
    p.Y = &H32323232
    p.Label = "POINTLBL"
    Set c = New Collection
    c.Add "one"
    c.Add "two"
    Call T_UdtByRef(p)
    Call T_UdtStrFirst(q)
    Call T_EnumArg(ColGreen)
    Call T_OptSupplied
    Call T_OptSupplied(&H41414141)
    Call T_OptVar
    Call T_ParamArr(1, 2, 3)
    Call T_CollArg(c)
End Sub
'@
     Invoke=@{ Name='T_UdtDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 8) {
            return "expected 8 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        # An Enum member IS a Long in VBA, so the Long load is the right
        # answer -- the enum's NAME is a compile-time construct and is simply
        # not present in the p-code.
        if ($entry['T_EnumArg'].args -ne 'a1:Long=20') {
            return "T_EnumArg args were [$($entry['T_EnumArg'].args)]" }
        # An omitted Optional is materialised by the CALLER, so its default is
        # visible as an ordinary value.
        if ($entry['T_OptSupplied'].args -ne 'a1:Long=77') {
            return "omitted Optional should show its default 77, got [$($entry['T_OptSupplied'].args)]" }
        # An omitted Optional VARIANT is the published `LitVar_Missing`:
        # VT_ERROR with DISP_E_PARAMNOTFOUND. Both constants are exact, which
        # is what makes it safe to decode with no type information at all.
        if ($entry['T_OptVar'].args -ne 'a1:?unseen=Missing') {
            return "omitted Optional Variant should read Missing, got [$($entry['T_OptVar'].args)]" }
        # A UDT is always by reference and the slot points straight at the
        # record. The TYPE is recovered; the record's field layout is not, so
        # the pointer is reported rather than invented contents.
        if ($entry['T_UdtByRef'].typetext -ne 'Udt&') {
            return "T_UdtByRef signature was [$($entry['T_UdtByRef'].typetext)]" }
        if ($entry['T_UdtByRef'].args -notmatch '^a1:Udt&=udt@0x[0-9A-F]+$') {
            return "T_UdtByRef args were [$($entry['T_UdtByRef'].args)]" }
        # Still the address when the first field is a String, never that field.
        if ($entry['T_UdtStrFirst'].typetext -ne 'Udt&') {
            return "T_UdtStrFirst signature was [$($entry['T_UdtStrFirst'].typetext)]" }
        if ($entry['T_UdtStrFirst'].args -notmatch '^a1:Udt&=udt@0x[0-9A-F]+$') {
            return "a record whose first field is a String rendered as [$($entry['T_UdtStrFirst'].args)]" }
        # A class reference is an Object.
        if ($entry['T_CollArg'].typetext -ne 'Object') {
            return "T_CollArg signature was [$($entry['T_CollArg'].typetext)]" }
        # A ParamArray arrives as a Variant SAFEARRAY and decodes as one.
        # The descriptor carries FADF_VARIANT rather than FADF_HAVEVARTYPE, and the
        # header and elements both take their type from EffectiveElemVt.
        if ($entry['T_ParamArr'].args -notmatch '^a1:Ref&=Variant\[0\.\.2\]') {
            return "T_ParamArr args were [$($entry['T_ParamArr'].args)]" }
        $null }
     Why='user-defined Types, Enums, class references, Optional and ParamArray
          -- what the decoder does with the constructs it has never seen' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
