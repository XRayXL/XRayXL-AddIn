# THE SAME PARAMETER LIST IN A CLASS AND A FORM.
#
# `frame-returns` proved the declared return type is invisible from the argument
# side. This proves the CONTAINER is too, and it is a separate claim: a class or
# form procedure is a COM method, so a Function returns its value through a
# TRAILING `[out, retval]` argument slot while the call itself returns an
# HRESULT. The walker counted that slot as a parameter, so every class and form
# Function reported `argcount` one too high and a spurious raw qword after the
# real arguments -- for EVERY return type, which is what separates it from the
# standard-module defect, where the extra slot is the FIRST and only `Variant`
# has one.
#
# Riding on the same blind spot: `ZeroRetVal` is in the exit-opcode family but
# does not end a procedure -- it zeroes the result at ENTRY, and the compiler
# emits it only for a return type that needs it (String, Variant, Object). The
# type walk stopped on it before reaching the first parameter load, so those
# three came back `(?,?,?)` with values rendered as raw qwords, while the same
# signature returning Long or Double decoded correctly.
#
# Neither was visible to 35,213 real-world signatures, because every one of them
# was emitted into a STANDARD module. 70% of the source they were mined from is
# class and form code.
$case = @{ Name='frame-returns-container'
     ClassSetup=@'
Public Function K_Var(ByVal a As Long, Optional b As Variant) As Variant
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    K_Var = 1
End Function
Public Function K_Long(ByVal a As Long, Optional b As Variant) As Long
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    K_Long = 2
End Function
Public Function K_Str(ByVal a As Long, Optional b As Variant) As String
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    K_Str = "S"
End Function
Public Function K_Obj(ByVal a As Long, Optional b As Variant) As Object
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    Set K_Obj = Nothing
End Function
Public Function K_Dbl(ByVal a As Long, Optional b As Variant) As Double
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    K_Dbl = 2748.5
End Function
Public Sub K_Sub(ByVal a As Long, Optional b As Variant)
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
End Sub
'@
     FormSetup=@'
Public Function U_Var(ByVal a As Long, Optional b As Variant) As Variant
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
    U_Var = 1
End Function
Public Sub U_Sub(ByVal a As Long, Optional b As Variant)
    Dim z As Long
    Dim v As Variant
    z = a
    v = b
End Sub
'@
     Setup=@'
Public Sub T_ContDrive()
    Dim k As New CCase
    Dim u As New UFCase
    Dim r As Variant
    Dim o As Object
    r = k.K_Var(&H11223344, 42)
    r = k.K_Var(&H11223344)
    r = k.K_Long(&H11223344, 42)
    r = k.K_Str(&H11223344, 42)
    Set o = k.K_Obj(&H11223344, 42)
    r = k.K_Dbl(&H11223344, 42)
    Call k.K_Sub(&H11223344, 42)
    Call k.K_Sub(&H11223344)
    r = u.U_Var(&H11223344, 42)
    Call u.U_Sub(&H11223344, 42)
End Sub
'@
     Invoke=@{ Name='T_ContDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        $e = @{}
        foreach ($r in $t.rows) {
            if ($r.kind -eq 'entry' -and $r.source -eq 'VBA' -and
                ($r.function -like 'K_*' -or $r.function -like 'U_*')) {
                if (-not $e.ContainsKey($r.function)) { $e[$r.function] = @() }
                $e[$r.function] += $r }
        }
        $all = @('K_Var','K_Long','K_Str','K_Obj','K_Dbl','K_Sub','U_Var','U_Sub')
        foreach ($fn in $all) { if (-not $e.ContainsKey($fn)) { return "$fn never entered" } }

        # ONE PARAMETER LIST, SIX RETURN TYPES, TWO CONTAINERS, ONE ANSWER --
        # and identical to what `frame-returns` asserts for a standard module.
        $supplied = 'a1:Long=287454020 a2:Variant&=Integer(42)'
        foreach ($fn in $all) {
            $r = $e[$fn][0]
            if ($r.argcount -ne '2') {
                return "$fn argcount was [$($r.argcount)], expected 2 -- a result slot counted as an argument" }
            if ($r.typetext -ne 'Long,Variant&') {
                return "$fn typetext was [$($r.typetext)], expected Long,Variant& -- the type walk stopped early" }
            if ($r.args -ne $supplied) {
                return "$fn args were [$($r.args)], expected [$supplied]" }
        }
        # An omitted Optional still reads Missing in a class module.
        $omitted = 'a1:Long=287454020 a2:Variant&=Missing'
        foreach ($fn in @('K_Var','K_Sub')) {
            if ($e[$fn].Count -lt 2) { return "$fn was entered $($e[$fn].Count) time(s), expected 2" }
            if ($e[$fn][1].args -ne $omitted) {
                return "$fn omitted-Optional args were [$($e[$fn][1].args)], expected [$omitted]" }
        }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='a class or form Function returns through a trailing argument slot and
          may zero its result on entry -- neither must reach the argument column' }

if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
