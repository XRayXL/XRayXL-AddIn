# The same parameters read identically behind five return types: a Variant-returning Function
# gets its result's address in the first slot, which is not an argument. An omitted Optional
# reads `Missing` even where the p-code types the slot.
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
        # T_RetVar and T_RetSub are each called twice, supplied then omitted
        $e = @{}
        foreach ($r in $t.rows) {
            if ($r.kind -eq 'entry' -and $r.source -eq 'VBA' -and $r.function -like 'T_Ret*') {
                if (-not $e.ContainsKey($r.function)) { $e[$r.function] = @() }
                $e[$r.function] += $r }
        }
        foreach ($fn in @('T_RetVar','T_RetLong','T_RetStr','T_RetObj','T_RetSub')) {
            if (-not $e.ContainsKey($fn)) { return "$fn never entered" }
        }
        # counting the result slot would read the Variant Function as (?,Long,Variant&)
        $supplied = 'a1:Long=287454020 a2:Variant&=Integer(42)'
        foreach ($fn in @('T_RetVar','T_RetLong','T_RetStr','T_RetObj','T_RetSub')) {
            $r = $e[$fn][0]
            if ($r.argcount -ne '2') {
                return "$fn argcount was [$($r.argcount)], expected 2 -- a result slot counted as an argument" }
            if ($r.typetext -ne 'Long,Variant&') {
                return "$fn typetext was [$($r.typetext)], expected Long,Variant&" }
            if ((Remove-ArgAddress $r.args) -ne $supplied) {
                return "$fn args were [$($r.args)], expected [$supplied]" }
        }
        # an omitted Optional reads `Missing`, including when the slot is typed
        $omitted = 'a1:Long=287454020 a2:Variant&=Missing'
        foreach ($fn in @('T_RetVar','T_RetSub')) {
            if ($e[$fn].Count -lt 2) { return "$fn was entered $($e[$fn].Count) time(s), expected 2" }
            if ((Remove-ArgAddress $e[$fn][1].args) -ne $omitted) {
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
