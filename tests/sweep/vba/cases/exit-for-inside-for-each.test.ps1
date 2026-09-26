# `Exit For` inside a `For Each` leaves the loop through its own opcode (1038-1042), which releases
# the enumerator and carries on in the procedure: one call, one row, its frame open past the loop,
# a call after the loop its child, and a Variant Function's parameters after its result buffer.
# `Exit Function` and `Exit Sub` inside a `For Each` leave through the procedure's own exit. VBA has
# no multi-level Exit For, so `GoTo`, `Exit Function` and `Exit Do` out of nested loops are here too.
$case = @{ Name='exit-for-inside-for-each'
     Setup=@'
Public gN As Long
Public Function T_EFColl(ByVal k As Long) As Variant
    Dim c As New Collection, x As Variant
    c.Add 1: c.Add 2: c.Add 3
    For Each x In c
        If x = 2 Then Exit For
    Next
    T_EFLeaf k
    T_EFColl = k
End Function
Public Sub T_EFArr(ByVal m As Long)
    Dim a(1 To 3) As Long, v As Variant, i As Long
    For Each v In a
        If i = 1 Then Exit For
        i = i + 1
    Next
    For i = 1 To 5
        If i = 2 Then Exit For
    Next
    T_EFLeaf m
End Sub
Public Sub T_EFVar(ByVal m As Long)
    Dim v As Variant, x As Variant
    v = Array(1, 2, 3)
    For Each x In v
        If x = 2 Then Exit For
    Next
    T_EFLeaf m
End Sub
Public Sub T_EFNested(ByVal m As Long)
    Dim c As New Collection, x As Variant, y As Variant
    c.Add 1: c.Add 2
    For Each x In c
        For Each y In c
            If y = 2 Then Exit For
        Next
        If x = 2 Then Exit For
    Next
    T_EFLeaf m
End Sub
Public Function T_EFExitFn(ByVal k As Long) As Long
    Dim a(1 To 3) As Long, v As Variant
    For Each v In a
        T_EFExitFn = k: Exit Function
    Next
    T_EFExitFn = -1
End Function
Public Sub T_EFExitSub(ByVal k As Long)
    Dim c As New Collection, x As Variant
    c.Add 1
    For Each x In c
        gN = k: Exit Sub
    Next
    gN = -1
End Sub
' VBA has no multi-level Exit For: GoTo, Exit Function or Exit Do leave several levels at once.
Public Sub T_EFGoTo(ByVal m As Long)
    Dim c As New Collection, x As Variant, y As Variant
    c.Add 1: c.Add 2
    For Each x In c
        For Each y In c
            If y = 2 Then GoTo Done
        Next
    Next
Done:
    T_EFLeaf m
End Sub
Public Function T_EFDeepExit(ByVal k As Long) As Variant
    Dim c As New Collection, x As Variant, y As Variant
    c.Add 1: c.Add 2
    For Each x In c
        For Each y In c
            If y = 2 Then T_EFDeepExit = k: Exit Function
        Next
    Next
    T_EFDeepExit = -1
End Function
Public Sub T_EFMixed(ByVal m As Long)
    Dim c As New Collection, x As Variant, i As Long, n As Long
    c.Add 1: c.Add 2
    For i = 1 To 3
        For Each x In c
            If x = 1 Then Exit For
        Next
        If i = 2 Then Exit For
    Next
    For Each x In c
        Do
            n = n + 1
            If n > 2 Then Exit Do
        Loop
    Next
    With c
        For Each x In c
            If x = 1 Then Exit For
        Next
    End With
    T_EFLeaf m
End Sub
Public Sub T_EFLeaf(ByVal k As Long)
    gN = k
End Sub
Public Sub T_EFDrive()
    Dim r As Variant
    r = T_EFColl(7)
    T_EFArr 9
    T_EFVar 10
    T_EFNested 11
    gN = T_EFExitFn(12)
    T_EFExitSub 13
    T_EFGoTo 14
    r = T_EFDeepExit(15)
    T_EFMixed 16
End Sub
'@
     Invoke=@{ Name='T_EFDrive'; Args=@() }
     Expect={ param($t)
        if ($t.maxDepth -ne 3) { return "expected depth 3, got $($t.maxDepth)" }
        $null }
     Calls=@(
        @{ Function='T_EFDrive';   Depth='1'; Parent=-1; Outcome='returned' }
        @{ Function='T_EFColl';    Depth='2'; Parent=0;  Args='a1:Long=7'; Ret='Long(7)'; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=1;  Args='a1:Long=7'; Outcome='returned' }
        @{ Function='T_EFArr';     Depth='2'; Parent=0;  Args='a1:Long=9'; Ret=''; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=3;  Args='a1:Long=9'; Outcome='returned' }
        @{ Function='T_EFVar';     Depth='2'; Parent=0;  Args='a1:Long=10'; Ret=''; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=5;  Args='a1:Long=10'; Outcome='returned' }
        @{ Function='T_EFNested';  Depth='2'; Parent=0;  Args='a1:Long=11'; Ret=''; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=7;  Args='a1:Long=11'; Outcome='returned' }
        @{ Function='T_EFExitFn';  Depth='2'; Parent=0;  Args='a1:Long=12'; Ret='12'; Outcome='returned' }
        @{ Function='T_EFExitSub'; Depth='2'; Parent=0;  Args='a1:Long=13'; Ret=''; Outcome='returned' }
        @{ Function='T_EFGoTo';    Depth='2'; Parent=0;  Args='a1:Long=14'; Ret=''; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=11; Args='a1:Long=14'; Outcome='returned' }
        @{ Function='T_EFDeepExit'; Depth='2'; Parent=0; Args='a1:Long=15'; Ret='Long(15)'; Outcome='returned' }
        @{ Function='T_EFMixed';   Depth='2'; Parent=0;  Args='a1:Long=16'; Ret=''; Outcome='returned' }
        @{ Function='T_EFLeaf';    Depth='3'; Parent=14; Args='a1:Long=16'; Outcome='returned' } )
     Why='a loop exit is not a procedure exit: one row per call, the frame open past the loop, the
          call after it a child, no return value for a Sub, and parameters numbered from the right slot' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
