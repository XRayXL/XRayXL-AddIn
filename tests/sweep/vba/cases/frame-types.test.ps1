# One parameter of every type. Types come from the typed load opcode, not the
# slot bytes, so each body reads its parameter; T_TyUnused covers one never read, which its caller types.
$case = @{ Name='frame-types'
     Setup=@'
Public Sub T_TyI2(ByVal v As Integer)
    Dim z As Integer
    z = v
End Sub
Public Sub T_TyI4(ByVal v As Long)
    Dim z As Long
    z = v
End Sub
Public Sub T_TyR4(ByVal v As Single)
    Dim z As Single
    z = v
End Sub
Public Sub T_TyR8(ByVal v As Double)
    Dim z As Double
    z = v
End Sub
Public Sub T_TyCy(ByVal v As Currency)
    Dim z As Currency
    z = v
End Sub
Public Sub T_TyStr(ByVal v As String)
    Dim z As String
    z = v
End Sub
Public Sub T_TyBool(ByVal v As Boolean)
    Dim z As Boolean
    z = v
End Sub
Public Sub T_TyDate(ByVal v As Date)
    Dim z As Date
    z = v
End Sub
Public Sub T_TyByte(ByVal v As Byte)
    Dim z As Byte
    z = v
End Sub
Public Sub T_TyVar(ByVal v As Variant)
    Dim z As Variant
    z = v
End Sub
Public Sub T_TyLL(ByVal v As LongLong)
    Dim z As LongLong
    z = v
End Sub
Public Sub T_TyObj(ByVal o As Object)
    Dim z As Object
    Set z = o
End Sub
Public Sub T_TyRefI4(ByRef v As Long)
    v = 7
End Sub
Public Sub T_TyRefStr(ByRef v As String)
    v = "REFSTR"
End Sub
Public Sub T_TyUnused(ByVal a As Long, ByVal b As Double)
    Dim z As Long
    z = 1
End Sub
Public Sub T_TyDrive()
    Dim o As Object
    Dim n As Long
    Dim s As String
    Dim ll As LongLong
    Set o = Application
    ll = 1234567890123^
    Call T_TyI2(1234)
    Call T_TyI4(&H11223344)
    Call T_TyR4(1.5)
    Call T_TyR8(2748.5)
    Call T_TyCy(9.99)
    Call T_TyStr("TYPESTR")
    Call T_TyBool(True)
    Call T_TyDate(#1/2/2020#)
    Call T_TyByte(200)
    Call T_TyVar(42)
    Call T_TyLL(ll)
    Call T_TyObj(o)
    n = 0
    Call T_TyRefI4(n)
    s = "X"
    Call T_TyRefStr(s)
    Call T_TyUnused(&H11223344, 2748.5)
End Sub
'@
     Invoke=@{ Name='T_TyDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.procedures -lt 16) {
            return "expected 16 procedures, got $($t.procedures)" }
        $e = Get-FirstEntryByName $t.rows
        # Knowing the type is what turns the bits back into the value.
        $wantArgs = @{
            'T_TyI2'='a1:Integer=1234'; 'T_TyI4'='a1:Long=287454020';
            'T_TyR4'='a1:Single=1.5';   'T_TyR8'='a1:Double=2748.5';
            'T_TyCy'='a1:Currency=9.9900'; 'T_TyStr'='a1:String="TYPESTR"';
            'T_TyBool'='a1:Integer=-1'; 'T_TyDate'='a1:Double=43832';
            'T_TyByte'='a1:Byte=200';   'T_TyLL'='a1:LongLong=1234567890123';
            'T_TyVar'='a1:Variant=Integer(42)' }
        foreach ($fn in $wantArgs.Keys) {
            if ($e[$fn].args -ne $wantArgs[$fn]) {
                return "$fn args were [$($e[$fn].args)], expected [$($wantArgs[$fn])]" }
        }
        # a parameter never read emits no load; its caller's literals type it instead
        if ($e['T_TyUnused'].typetext -ne 'Long,Double' -or $e['T_TyUnused'].args -ne 'a1:Long=287454020 a2:Double=2748.5') {
            return "T_TyUnused should read its caller's Long and Double, got [$($e['T_TyUnused'].typetext)] [$($e['T_TyUnused'].args)]" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='one parameter of every VBA type, so each typed load opcode can be
          read out of the p-code and mapped to the declared type' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
