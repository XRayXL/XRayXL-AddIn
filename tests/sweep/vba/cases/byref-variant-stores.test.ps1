# A write-only `ByRef Variant` is typed from whichever of its three stores the right-hand side
# picks (774 number, 783 Set, 787 full copy); store = load + 32 reaches only 774, since the other
# two sit above loads that name no type.
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
        # written out, not derived, so the test cannot just agree with the tracer
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
            # exact, so a ByVal reported as ByRef fails
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
