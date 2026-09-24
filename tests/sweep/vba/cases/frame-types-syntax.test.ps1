# Type-suffix spellings (`v%`, `v&`, `v$`...) must decode exactly as their `As` forms.
# Also: bare untyped parameters, suffix arrays, ParamArray after a positional
# parameter (where indexing can slip), and a library-qualified class.
$case = @{ Name='frame-types-syntax'
     Setup=@'
Public Sub T_SfxInt(ByVal v%)
    Dim z As Integer
    z = v
End Sub
Public Sub T_SfxLng(ByVal v&)
    Dim z As Long
    z = v
End Sub
Public Sub T_SfxLL(ByVal v^)
    Dim z As LongLong
    z = v
End Sub
Public Sub T_SfxSng(ByVal v!)
    Dim z As Single
    z = v
End Sub
Public Sub T_SfxDbl(ByVal v#)
    Dim z As Double
    z = v
End Sub
Public Sub T_SfxCur(ByVal v@)
    Dim z As Currency
    z = v
End Sub
Public Sub T_SfxStr(ByVal v$)
    Dim z As String
    z = v
End Sub
Public Sub T_BareVar(ByVal v)
    Dim z As Variant
    z = v
End Sub
Public Sub T_SfxArr(ByRef a#())
    Dim z As Double
    z = a(1)
End Sub
Public Sub T_PtrByVal(ByVal v As LongPtr)
    Dim z As LongPtr
    z = v
End Sub
Public Sub T_PaAfter(ByVal a As Long, ParamArray rest() As Variant)
    Dim z As Long
    z = a
    z = UBound(rest)
End Sub
Public Sub T_QualRange(ByVal r As Excel.Range)
    Dim z As Long
    z = r.Row
End Sub
Public Sub T_SynDrive()
    Dim d(1 To 3) As Double
    Dim ll As LongLong
    Dim pp As LongPtr
    d(1) = 4.25
    d(2) = 8.5
    ll = 1234567890123^
    pp = 8192
    Call T_SfxInt(1234)
    Call T_SfxLng(&H11223344)
    Call T_SfxLL(ll)
    Call T_SfxSng(1.5)
    Call T_SfxDbl(2748.5)
    Call T_SfxCur(9.99)
    Call T_SfxStr("SUFFIXSTR")
    Call T_BareVar(42)
    Call T_SfxArr(d)
    Call T_PtrByVal(pp)
    Call T_PaAfter(&H55667788, 1, 2, 3)
    Call T_QualRange(ThisWorkbook.Worksheets(1).Range("A1"))
End Sub
'@
     Invoke=@{ Name='T_SynDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 13) {
            return "expected 13 procedures, got $($t.procedures)" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $e = Get-FirstEntryByName $t.rows
        $wantSig = @{
            'T_SfxInt'='Integer'; 'T_SfxLng'='Long';
            'T_SfxLL'='LongLong'; 'T_SfxSng'='Single';
            'T_SfxDbl'='Double';  'T_SfxCur'='Currency';
            'T_SfxStr'='String';
            # no `As`, no suffix, no Def directive in this module: Variant
            'T_BareVar'='Variant';
            # LongPtr is an alias for LongLong on x64; the p-code cannot say "LongPtr"
            'T_PtrByVal'='LongLong';
            'T_QualRange'='Object' }
        foreach ($fn in $wantSig.Keys) {
            if (-not $e.ContainsKey($fn)) { return "$fn was not traced" }
            if ($e[$fn].typetext -ne $wantSig[$fn]) {
                return "$fn signature was [$($e[$fn].typetext)], expected $($wantSig[$fn])" }
        }
        # a suffix decoded to the right name but the wrong width would pass the signature check
        $wantArgs = @{
            'T_SfxInt'='a1:Integer=1234'; 'T_SfxLng'='a1:Long=287454020';
            'T_SfxSng'='a1:Single=1.5';   'T_SfxDbl'='a1:Double=2748.5';
            'T_SfxCur'='a1:Currency=9.9900';
            'T_SfxStr'='a1:String="SUFFIXSTR"';
            'T_SfxLL'='a1:LongLong=1234567890123' }
        foreach ($fn in $wantArgs.Keys) {
            if ($e[$fn].args -ne $wantArgs[$fn]) {
                return "$fn args were [$($e[$fn].args)], expected [$($wantArgs[$fn])]" }
        }
        # a bare untyped parameter is a ByVal Variant: three slots, one parameter
        if ([int]$e['T_BareVar'].argcount -ne 1) {
            return "T_BareVar argcount=$($e['T_BareVar'].argcount), expected 1 (ByVal Variant is 3 slots)" }
        if ($e['T_SfxArr'].args -notmatch '^a1:Ref&=Double\[1\.\.3\]') {
            return "T_SfxArr args were [$($e['T_SfxArr'].args)]" }
        # the ParamArray does not start at slot 1, so argument indexing can slip
        if ($e['T_PaAfter'].args -notmatch '^a1:Long=1432778632') {
            return "T_PaAfter first arg was [$($e['T_PaAfter'].args)], expected a1:Long=1432778632 leading" }
        if ([int]$e['T_PaAfter'].argcount -ne 2) {
            return "T_PaAfter argcount=$($e['T_PaAfter'].argcount), expected 2" }
        $null }
     Why='the type-suffix spellings, the bare untyped Variant, a suffix-form
          array, a ParamArray after a positional parameter, and a qualified
          class name -- the parameter spellings MS-VBAL admits that the
          As-clause cases do not reach' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
