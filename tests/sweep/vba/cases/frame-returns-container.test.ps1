# The same parameters read identically in a class and a form, whose Functions return through a
# trailing `[out, retval]` slot that is not a parameter. `ZeroRetVal` is in the exit-opcode
# family but ends nothing, so the type walk must not stop on it.
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

        # identical to what `frame-returns` asserts for a standard module
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
