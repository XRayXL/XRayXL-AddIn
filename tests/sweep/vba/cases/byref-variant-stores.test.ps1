# A `ByRef Variant` that is only written has three store opcodes, not one.
#
# A write-only parameter emits no load, so its type comes from the store. store = load + 32
# reaches 774 = 742+32; the other two ByRef Variant stores, 783 and 787, sit above opcodes that
# name no type.
#
# The right-hand side chooses which is emitted: a number takes 774, `Set` takes 783, and
# anything needing a full copy (String, Variant, array) takes 787. The cases vary the parameter
# type under each of those shapes: only a `ByRef Variant` produces 783 or 787, and every `ByVal
# Variant` produces 1477.
$case = @{ Name='byref-variant-stores'
     Setup=@'
' --- the three ByRef Variant stores, one per right-hand side --------------
Public Sub S_VarNum(ByRef v As Variant)
    v = 42#
End Sub
Public Sub S_VarSet(ByRef v As Variant)
    Set v = Application
End Sub
Public Sub S_VarStr(ByRef v As Variant)
    v = "text"
End Sub
Public Sub S_VarArr(ByRef v As Variant)
    Dim t(0 To 1) As Double
    t(1) = 1#
    v = t
End Sub

' --- the SAME assignments into other ByRef types: each has its own store ---
Public Sub S_ObjSet(ByRef o As Object)
    Set o = Application
End Sub
Public Sub S_ClsSet(ByRef c As Collection)
    Set c = New Collection
End Sub
Public Sub S_StrStr(ByRef s As String)
    s = "text"
End Sub
Public Sub S_LngNum(ByRef n As Long)
    n = 42
End Sub
Public Sub S_DblNum(ByRef d As Double)
    d = 42#
End Sub

' --- ByVal Variant, which must stay ByVal whatever is assigned ------------
Public Sub S_ByValStr(ByVal v As Variant)
    v = "text"
End Sub
Public Sub S_ByValSet(ByVal v As Variant)
    Set v = Application
End Sub

Public Sub S_StoreDrive()
    Dim v As Variant, o As Object, c As Collection, s As String
    Dim n As Long, d As Double
    S_VarNum v
    v = 1#: S_VarSet v
    v = 1#: S_VarStr v
    v = 1#: S_VarArr v
    S_ObjSet o
    S_ClsSet c
    S_StrStr s
    S_LngNum n
    S_DblNum d
    v = 1#: S_ByValStr v
    v = 1#: S_ByValSet v
End Sub
'@
     Invoke=@{ Name='S_StoreDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        # The expected text is written out per procedure rather than derived, so
        # the test knows the answer it is asserting rather than agreeing with
        # whatever the tracer produced.
        $want = @{
            'S_VarNum'   = 'Variant&'
            'S_VarSet'   = 'Variant&'
            'S_VarStr'   = 'Variant&'
            'S_VarArr'   = 'Variant&'
            'S_ObjSet'   = 'Object&'
            'S_ClsSet'   = 'Object&'
            'S_StrStr'   = 'String&'
            'S_LngNum'   = 'Long&'
            'S_DblNum'   = 'Double&'
            'S_ByValStr' = 'Variant'
            'S_ByValSet' = 'Variant'
        }
        $bad = @()
        foreach ($fn in $want.Keys) {
            if (-not $entry.ContainsKey($fn)) { $bad += "$fn=MISSING"; continue }
            $got = [string]$entry[$fn].typetext
            # Exact, so a ByVal reported as ByRef fails: `Variant` and `Variant&`
            # are different answers and the & is the whole point of the family.
            if ($got -ne $want[$fn]) { $bad += "$fn=$got (want $($want[$fn]))" }
        }
        if ($bad.Count -gt 0) { return "wrong store type: $($bad -join '; ')" }
        $null }
     Why='a write-only `ByRef Variant` read `?op787` (or `?op783`) in a real
          workbook: two of its three store opcodes sit above loads that carry no
          type, so store=load+32 cannot reach them' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
