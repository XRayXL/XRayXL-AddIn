# An early `Exit` must not truncate the signature, in any kind of procedure.
#
# `Exit Sub`, `Exit Function` and `Exit Property` emit the same opcode as the real end of the
# procedure, so a walk that stops at the first terminator leaves every parameter first used
# after it as `?unseen`.
#
# The exit opcode is chosen by return type: a Sub leaves by 635, a Double or Date by 627, Long
# 625, Boolean 624, String 630, Object 631, an array 634, Variant 952, and a class member by
# 1664 when it returns a value or 504 when it does not. Hence a table of return types.
#
# Every procedure takes `a` (tested before the exit) and `b` (used only after it), so a
# truncated walk shows up as `b` unresolved while `a` is typed.
$case = @{ Name='early-exit-types'
     ClassSetup=@'
Public Val As Double

' A class member that RETURNS -- leaves by a different exit from one that does not
Public Function CFun(ByVal a As Double, ByVal b As Double) As Double
    If a < -999 Then
        CFun = 0
        Exit Function
    End If
    CFun = b
End Function
Public Sub CSub(ByVal a As Double, ByVal b As Double)
    If a < -999 Then Exit Sub
    Val = b
End Sub
Public Property Get CProp(ByVal a As Double) As Double
    If a < -999 Then Exit Property
    CProp = a
End Property
Public Property Let CPropLet(ByVal a As Double, ByVal b As Double)
    If a < -999 Then Exit Property
    Val = b
End Property
'@
     Setup=@'
Private mM As Double

' --- one per return type, since the return type picks the exit opcode ------
Public Sub X_Sub(ByVal a As Double, ByVal b As Double)
    If a < -999 Then Exit Sub
    Dim z As Double
    z = b
End Sub
Public Function X_Dbl(ByVal a As Double, ByVal b As Double) As Double
    If a < -999 Then Exit Function
    X_Dbl = b
End Function
Public Function X_Lng(ByVal a As Double, ByVal b As Long) As Long
    If a < -999 Then Exit Function
    X_Lng = b
End Function
Public Function X_Bool(ByVal a As Double, ByVal b As Boolean) As Boolean
    If a < -999 Then Exit Function
    X_Bool = b
End Function
Public Function X_Str(ByVal a As Double, ByVal b As String) As String
    If a < -999 Then Exit Function
    X_Str = b
End Function
Public Function X_Var(ByVal a As Double, ByVal b As Variant) As Variant
    If a < -999 Then Exit Function
    X_Var = b
End Function
Public Function X_Obj(ByVal a As Double, ByVal b As Object) As Object
    If a < -999 Then Exit Function
    Set X_Obj = b
End Function
Public Function X_Date(ByVal a As Double, ByVal b As Date) As Date
    If a < -999 Then Exit Function
    X_Date = b
End Function
Public Function X_Arr(ByVal a As Double, ByVal b As Double) As Double()
    If a < -999 Then Exit Function
    Dim t(0 To 1) As Double
    t(1) = b
    X_Arr = t
End Function

' --- module-level Property accessors, both directions ---------------------
Public Property Get X_Prop(ByVal a As Double) As Double
    If a < -999 Then Exit Property
    X_Prop = a
End Property
Public Property Let X_Prop(ByVal a As Double, ByVal b As Double)
    If a < -999 Then Exit Property
    mM = b
End Property

' --- the shapes an exit can take ------------------------------------------
' Colon-joined to the assignment, rather than on its own line.
Public Function X_Colon(ByVal a As Double, ByVal b As Double) As Double
    If a < -999 Then X_Colon = 0: Exit Function
    X_Colon = b
End Function
' More than one early exit, so the walk must pass both.
Public Sub X_Twice(ByVal a As Double, ByVal b As Double)
    If a < -999 Then Exit Sub
    If a < -998 Then Exit Sub
    Dim z As Double
    z = b
End Sub
' The LAST statement is itself an Exit: nothing follows, so the walk must stop
' there rather than step off the end of the procedure.
Public Function X_EndsOnExit(ByVal a As Double, ByVal b As Double) As Double
    X_EndsOnExit = b
    If a < -999 Then Exit Function
End Function

Public Sub X_ExitDrive()
    Dim d As Double, s As String, v As Variant, l As Long, bo As Boolean
    Dim ob As Object, ar() As Double, dt As Date
    X_Sub 1#, 2#
    d = X_Dbl(1#, 2#)
    l = X_Lng(1#, 2&)
    bo = X_Bool(1#, True)
    s = X_Str(1#, "x")
    v = X_Var(1#, 2#)
    Set ob = X_Obj(1#, Application)
    dt = X_Date(1#, #1/2/2020#)
    ar = X_Arr(1#, 2#)
    X_Prop(1#) = 2#
    d = X_Prop(1#)
    d = X_Colon(1#, 2#)
    X_Twice 1#, 2#
    d = X_EndsOnExit(1#, 2#)
    Dim c As New CCase
    d = c.CFun(1#, 2#)
    c.CSub 1#, 2#
    d = c.CProp(1#)
    c.CPropLet(1#) = 2#
End Sub
'@
     Invoke=@{ Name='X_ExitDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $entry = Get-FirstEntryByName $t.rows
        # `b` is the LAST parameter of every one of these, and it is used only
        # after the early exit. If the walk stopped at that exit, `b` is the one
        # that goes unresolved -- so the last position carrying a '?' is the
        # signal, and naming it makes a regression say which shape broke.
        $truncated = @(); $missing = @()
        foreach ($fn in @('X_Sub','X_Dbl','X_Lng','X_Bool','X_Str','X_Var',
                          'X_Obj','X_Date','X_Arr','X_Prop','X_Colon',
                          'X_Twice','X_EndsOnExit',
                          'CFun','CSub','CProp','CPropLet')) {
            if (-not $entry.ContainsKey($fn)) { $missing += $fn; continue }
            $tt = [string]$entry[$fn].typetext
            # A '?' at all, or the `~` that says the walk resynchronised.
            if ($tt -match '\?' -or $tt -match '~') { $truncated += "$fn=$tt" }
        }
        if ($missing.Count -gt 0) { return "never traced: $($missing -join ' ')" }
        if ($truncated.Count -gt 0) {
            return "early Exit truncated the walk: $($truncated -join ' ')" }
        $null }
     Why='an early Exit emits the procedure''s terminator, so a walk that stops
          on one stops at the first early return and leaves everything after it
          `?unseen` -- 6,335 rows of it in a real workbook. The exit opcode is
          chosen by return type, so every return type is here.' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
