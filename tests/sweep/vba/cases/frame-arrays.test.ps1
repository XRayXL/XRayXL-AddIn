# Arrays with known bounds and contents. Bounds differ per array so a SAFEARRAY
# read at the wrong offset gives obviously wrong numbers.
$case = @{ Name='frame-arrays'
     Setup=@'
Public Sub T_ArrL(ByRef a() As Long)
    Dim z As Long
    z = a(1)
End Sub
Public Sub T_ArrS(ByRef s() As String)
    Dim z As String
    z = s(1)
End Sub
Public Sub T_Arr2(ByRef m() As Long)
    Dim z As Long
    z = m(1, 1)
End Sub
Public Sub T_ArrV(ByVal v As Variant)
    Dim z As Long
    z = 1
End Sub
Public Sub T_ArrDrive()
    Dim a(1 To 5) As Long
    Dim s(1 To 3) As String
    Dim m(1 To 2, 1 To 3) As Long
    Dim i As Long
    For i = 1 To 5
        a(i) = &H1000 + i
    Next i
    s(1) = "ALPHA"
    s(2) = "BETA"
    s(3) = "GAMMA"
    m(1, 1) = &H2001
    m(2, 3) = &H2006
    Call T_ArrL(a)
    Call T_ArrS(s)
    Call T_Arr2(m)
    Call T_ArrV(a)
End Sub
'@
     Invoke=@{ Name='T_ArrDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 5) {
            return "expected 5 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        foreach ($fn in @('T_ArrL','T_ArrS','T_Arr2','T_ArrV')) {
            if (-not $entry.ContainsKey($fn)) { return "no VBA entry row for $fn" }
        }
        # argcount counts parameters, not slots: a ByRef array is one slot
        foreach ($fn in @('T_ArrL','T_ArrS','T_Arr2')) {
            if ([int]$entry[$fn].argcount -ne 1) {
                return "$fn slots=$($entry[$fn].argcount), expected 1" }
        }
        # a ByVal Variant takes three slots but is still one parameter
        if ([int]$entry['T_ArrV'].argcount -ne 1) {
            return "T_ArrV argcount=$($entry['T_ArrV'].argcount), expected 1 (ByVal Variant is 3 slots, 1 parameter)" }

        # decimal, not raw bytes: the args column shares the result column's renderer
        if ((Remove-ArgAddress $entry['T_ArrL'].args) -ne 'a1:Ref&=Long[1..5]{4097,4098,4099,4100,4101}') {
            return "T_ArrL args were [$($entry['T_ArrL'].args)]" }
        if ((Remove-ArgAddress $entry['T_ArrS'].args) -ne 'a1:Ref&=String[1..3]{"ALPHA","BETA","GAMMA"}') {
            return "T_ArrS args were [$($entry['T_ArrS'].args)]" }
        # 2-D, and rendered in VBA declaration order: m(1 To 2, 1 To 3).
        if ((Remove-ArgAddress $entry['T_Arr2'].args) -notmatch '^a1:Ref&=Long\[1\.\.2,1\.\.3\]') {
            return "T_Arr2 args were [$($entry['T_Arr2'].args)]" }
        # the p-code names the slot Variant, though it carries VT_ARRAY|VT_I4
        if ($entry['T_ArrV'].typetext -ne 'Variant') {
            return "T_ArrV signature was [$($entry['T_ArrV'].typetext)]" }
        if ($entry['T_ArrV'].args -notmatch '^a1:Variant=') {
            return "T_ArrV args were [$($entry['T_ArrV'].args)]" }
        $null }
     Why='establishes how an array argument reaches the frame, with bounds and
          contents known in advance' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
