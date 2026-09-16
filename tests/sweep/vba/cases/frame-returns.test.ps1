# WHAT THE RETURN TYPE DOES TO THE ARGUMENTS.
#
# Every other parameter case in this suite is a `Sub`, and that is why two
# defects lived here for months. A Function returning `Variant` is handed the
# address of the caller's result VARIANT in an argument slot -- the calling
# convention pushes it last, so it arrives first -- and the walker counted it as
# argument 1: `argcount` one too high, a leading `?` in the signature, and every
# real argument reported one place to the right. It is not an exotic shape. It
# is what a UDF is.
#
# The second defect rode on the same blind spot. An omitted `Optional` IS a
# VARIANT carrying VT_ERROR / DISP_E_PARAMNOTFOUND, so when the p-code declared
# the slot a Variant the Variant decoder rendered the marker's payload,
# `Error(0x80020004)`, instead of the `Missing` the row model promises. The
# existing case only ever asked it of a slot the p-code left untyped.
#
# So this asserts the SAME parameter list behind three different return types.
# The arguments must read identically in all three: the return type is not
# supposed to be visible from the argument side at all.
$case = @{ Name='frame-returns'
     Setup=@'
Public Function T_RetVar(ByVal a As Long, Optional b As Variant) As Variant
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    T_RetVar = 1
End Function
Public Function T_RetLong(ByVal a As Long, Optional b As Variant) As Long
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    T_RetLong = 2
End Function
Public Function T_RetStr(ByVal a As Long, Optional b As Variant) As String
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    T_RetStr = "S"
End Function
Public Function T_RetObj(ByVal a As Long, Optional b As Variant) As Object
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    Set T_RetObj = Nothing
End Function
Public Sub T_RetSub(ByVal a As Long, Optional b As Variant)
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
End Sub
Public Sub T_RetDrive()
    Dim r As Variant
    Dim o As Object
    r = T_RetVar(&H11223344, 42)
    r = T_RetVar(&H11223344)
    r = T_RetLong(&H11223344, 42)
    r = T_RetStr(&H11223344, 42)
    Set o = T_RetObj(&H11223344, 42)
    Call T_RetSub(&H11223344, 42)
    Call T_RetSub(&H11223344)
End Sub
'@
     Invoke=@{ Name='T_RetDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        # Entry rows per procedure, in order -- T_RetVar and T_RetSub are each
        # called twice, supplied then omitted, and both calls are asserted.
        $e = @{}
        foreach ($r in $t.rows) {
            if ($r.kind -eq 'entry' -and $r.source -eq 'VBA' -and $r.function -like 'T_Ret*') {
                if (-not $e.ContainsKey($r.function)) { $e[$r.function] = @() }
                $e[$r.function] += $r }
        }
        foreach ($fn in @('T_RetVar','T_RetLong','T_RetStr','T_RetObj','T_RetSub')) {
            if (-not $e.ContainsKey($fn)) { return "$fn never entered" }
        }
        # ONE PARAMETER LIST, FIVE RETURN TYPES, ONE ANSWER. The Variant-
        # returning Function is the one that misreads as (?,Long,Variant&) with
        # argcount 3 when the result slot is counted as a parameter.
        $supplied = 'a1:Long=287454020 a2:Variant&=Integer(42)'
        foreach ($fn in @('T_RetVar','T_RetLong','T_RetStr','T_RetObj','T_RetSub')) {
            $r = $e[$fn][0]
            if ($r.argcount -ne '2') {
                return "$fn argcount was [$($r.argcount)], expected 2 -- a result slot counted as an argument" }
            if ($r.typetext -ne 'Long,Variant&') {
                return "$fn typetext was [$($r.typetext)], expected Long,Variant&" }
            if ($r.args -ne $supplied) {
                return "$fn args were [$($r.args)], expected [$supplied]" }
        }
        # AN OMITTED Optional READS `Missing`, INCLUDING WHEN THE SLOT IS TYPED.
        $omitted = 'a1:Long=287454020 a2:Variant&=Missing'
        foreach ($fn in @('T_RetVar','T_RetSub')) {
            if ($e[$fn].Count -lt 2) { return "$fn was entered $($e[$fn].Count) time(s), expected 2" }
            if ($e[$fn][1].args -ne $omitted) {
                return "$fn omitted-Optional args were [$($e[$fn][1].args)], expected [$omitted]" }
        }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='the declared return type changes the frame -- a Variant result occupies
          an argument slot -- and must not change what the arguments read as' }

if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
