# Planted answers for the frame layout. Argument count varies 0..3 so the
# argument-size field (bytes, not a count) shows itself; values are distinctive
# positive Longs so a slot is identified, not just consistent.
# No expression may overflow: an unhandled VBA error from Application.Run is a modal dialog.
$case = @{ Name='frame-layout'
     Setup=@'
Public Sub T_FL0()
    Dim z As Long
    z = 1
End Sub
Public Sub T_FL1(ByVal a As Long)
    Dim z As Long
    z = a
End Sub
Public Sub T_FL2(ByVal a As Long, ByVal b As Long)
    Dim z As Long
    z = a
    z = b
End Sub
Public Sub T_FL3(ByVal a As Long, ByVal b As Long, ByVal c As Long)
    Dim z As Long
    z = a
    z = b
    z = c
End Sub
Public Function T_FLRet(ByVal a As Long) As Long
    T_FLRet = &H0BADF00D
End Function
Public Function T_FLRetD(ByVal a As Long) As Double
    T_FLRetD = 1234.5
End Function
Public Function T_FLRetS(ByVal a As Long) As String
    T_FLRetS = "RETSENTINEL"
End Function
Public Function T_FLRetZ() As Long
    T_FLRetZ = &H7E7E7E7E
End Function
Public Function T_FLRetNone(ByVal a As Long) As Long
    Dim q As Long
    q = a
End Function
Public Sub T_FLStr(ByVal s As String, ByVal n As Long)
    Dim z As Long
    z = Len(s)
End Sub
Public Sub T_FLRef(ByRef a As Long, ByRef s As String)
    a = &H3B3B3B3B
    s = "REFOUT"
End Sub
Public Sub T_FLDbl(ByVal d As Double, ByVal a As Long)
    Dim z As Double
    z = d
End Sub
Public Sub T_FLDrive()
    Dim r As Long
    Dim rd As Double
    Dim ra As Long
    Dim rs As String
    Dim rs2 As String
    Dim rz As Long
    Dim rn As Long
    Call T_FL0
    Call T_FL1(&H11223344)
    Call T_FL2(&H11223344, &H55667788)
    Call T_FL3(&H11223344, &H55667788, &H1A2B3C4D)
    r = T_FLRet(&H11223344)
    rd = T_FLRetD(&H11223344)
    rs2 = T_FLRetS(&H11223344)
    rz = T_FLRetZ()
    rn = T_FLRetNone(&H11223344)
    Call T_FLStr("XRAYSENTINEL", &H2A2A2A2A)
    ra = &H4C4C4C4C
    rs = "BEFORE"
    Call T_FLRef(ra, rs)
    Call T_FLDbl(2748.5, &H5D5D5D5D)
End Sub
'@
     Invoke=@{ Name='T_FLDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        # T_FLDrive plus the twelve it calls.
        if ($t.procedures -lt 13) {
            return "expected 13 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }

        # The decoded arguments, asserted against the signatures written above.
        # This is the part that would have caught reading r15 as the frame
        # base: totals looked perfect while every argument was rubbish.
        $entry = Get-FirstEntryByName $t.rows
        foreach ($fn in @('T_FL0','T_FL1','T_FL2','T_FL3','T_FLStr','T_FLDbl','T_FLRef')) {
            if (-not $entry.ContainsKey($fn)) { return "no VBA entry row for $fn" }
        }
        # Argument COUNT comes from argSz, which is the whole point of it.
        $wantCount = @{ 'T_FL0'=0; 'T_FL1'=1; 'T_FL2'=2; 'T_FL3'=3;
                        'T_FLStr'=2; 'T_FLDbl'=2; 'T_FLRef'=2 }
        foreach ($fn in $wantCount.Keys) {
            $got = [int]$entry[$fn].argcount
            if ($got -ne $wantCount[$fn]) {
                return "$fn argcount=$got, expected $($wantCount[$fn])" }
        }
        # The SIGNATURE, recovered from each procedure's own p-code. This is
        # the assertion that the typed-opcode walk stayed synchronised: a
        # desynced walk yields plausible types for the wrong positions.
        $wantSig = @{ 'T_FL1'='Long'; 'T_FL2'='Long,Long';
                      'T_FL3'='Long,Long,Long'; 'T_FLStr'='String,?unseen';
                      'T_FLDbl'='Double,?unseen' }
        # T_FLRef's second slot reads String& or ? from run to run: both are honest, a wrong type never is.
        if ($entry['T_FLRef'].typetext -notin @('Long&,String&','Long&,?')) {
            return "T_FLRef signature was [$($entry['T_FLRef'].typetext)], expected Long&,String& or Long&,?" }
        foreach ($fn in $wantSig.Keys) {
            if ($entry[$fn].typetext -ne $wantSig[$fn]) {
                return "$fn signature was [$($entry[$fn].typetext)], expected $($wantSig[$fn])" }
        }
        # A '?unseen' is a parameter the body never READS: no opcode touches
        # its slot at all, so nothing was observed and there is genuinely no
        # type to recover. Asserting it keeps that honest -- the tracer must say
        # it does not know, and say WHICH not-knowing, rather than invent one.
        if ($entry['T_FLStr'].typetext -notmatch '\?\w*$') {
            return "T_FLStr's unread second parameter should be unknown" }

        # Argument VALUES, rendered USING the recovered type.
        if ($entry['T_FL3'].args -ne 'a1:Long=287454020 a2:Long=1432778632 a3:Long=439041101') {
            return "T_FL3 args were [$($entry['T_FL3'].args)]" }
        if ($entry['T_FL1'].args -ne 'a1:Long=287454020') {
            return "T_FL1 args were [$($entry['T_FL1'].args)]" }
        if ($entry['T_FL0'].args -ne '') {
            return "T_FL0 takes no arguments but reported [$($entry['T_FL0'].args)]" }
        # 2748.5 is exactly 0x40A5790000000000 as an IEEE double; knowing the
        # type is what turns those bits back into the number that was passed.
        if ($entry['T_FLDbl'].args -ne 'a1:Double=2748.5 a2:?unseen=0x5D5D5D5D') {
            return "T_FLDbl args were [$($entry['T_FLDbl'].args)]" }
        if ($entry['T_FLStr'].args -ne 'a1:String="XRAYSENTINEL" a2:?unseen=0x2A2A2A2A') {
            return "T_FLStr args were [$($entry['T_FLStr'].args)]" }
        # A ByRef slot holds a POINTER; knowing it is `Long&` is what licenses
        # following it. 0x4C4C4C4C is what the caller had assigned.
        if ($entry['T_FLRef'].args -notmatch '^a1:Long&=1280068684 ') {
            return "T_FLRef ByRef arg was [$($entry['T_FLRef'].args)]" }
        $null }
     Why='carries the known signatures every other frame case is read against,
          and asserts the decoded argument counts, types and values against
          them' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
