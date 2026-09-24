# Real-world parameter lists, mined from published VBA and chosen to cover what changes the
# frame (return type, passing mode, declared type, Optional supplied or omitted), read back
# with planted values. Committed, so the suite needs no add-ins and no network.
$case = @{ Name='frame-wild'
     Setup=@'
Public gaBool() As Boolean
Public gaByte() As Byte
Public gaCur() As Currency
Public gaDate() As Date
Public gaDouble() As Double
Public gaInt() As Integer
Public gaLL() As LongLong
Public gaLong() As Long
Public gaObj() As Object
Public gaSingle() As Single
Public gaString() As String
Public gaVar() As Variant
Public Sub W_Init()
    ReDim gaBool(0 To 2)
    ReDim gaByte(0 To 2)
    ReDim gaCur(0 To 2)
    ReDim gaDate(0 To 2)
    ReDim gaDouble(0 To 2)
    ReDim gaInt(0 To 2)
    ReDim gaLL(0 To 2)
    ReDim gaLong(0 To 2)
    ReDim gaObj(0 To 2)
    ReDim gaSingle(0 To 2)
    ReDim gaString(0 To 2)
    ReDim gaVar(0 To 2)
End Sub

Public Function W_0401(ByVal p1 As Object, ByVal p2 As LongPtr, p3 As LongPtr, p4 As Long) As LongPtr
    Dim o1 As Object
    Dim v2 As LongPtr
    Dim v3 As LongPtr
    Dim v4 As Long
    Set o1 = p1
    v2 = p2
    v3 = p3
    v4 = p4
    W_0401 = 1234567890123^
End Function

Public Function W_0442(ByRef p1() As Single, Optional ByVal p2 As Long = &H11223344, Optional ByVal p3 As Long = &H11223344, Optional ByVal p4 As Long = &H11223344) As Long
    Dim v2 As Long
    Dim v3 As Long
    Dim v4 As Long
    v2 = p2
    v3 = p3
    v4 = p4
    W_0442 = &H11223344
End Function

Public Function W_0448(ByVal p1 As Double) As Byte
    Dim v1 As Double
    v1 = p1
    W_0448 = 200
End Function

Public Function W_0951(ByRef p1 As Currency) As Currency
    Dim v1 As Currency
    v1 = p1
    W_0951 = 9.99@
End Function

Public Function W_1023(ByVal p1 As LongLong) As LongLong
    Dim v1 As LongLong
    v1 = p1
    W_1023 = 1234567890123^
End Function

Public Sub W_1027(ByVal p1 As LongPtr, ByVal p2 As Currency)
    Dim v1 As LongPtr
    Dim v2 As Currency
    v1 = p1
    v2 = p2
End Sub

Public Function W_10595(p1 As Object, ByVal p2 As Variant, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long, ByVal p6 As Long, Optional ByVal p7 As Long, Optional ByVal p8 As Long, Optional p9 As Long, Optional p10 As Long, Optional p11 As Long = &H11223344, Optional ByVal p12 As Object, Optional p13 As String = "WILD", Optional p14 As Byte, Optional p15 As Byte = 200, Optional p16 As Boolean, Optional p17 As Single = 1.5, Optional ByVal p18 As Single, Optional ByVal p19 As Boolean, Optional ByVal p20 As Byte = 200, Optional ByVal p21 As Boolean = True) As Boolean
    Dim o1 As Object
    Dim v2 As Variant
    Dim v3 As Long
    Dim v4 As Long
    Dim v5 As Long
    Dim v6 As Long
    Dim v7 As Long
    Dim v8 As Long
    Dim v9 As Long
    Dim v10 As Long
    Dim v11 As Long
    Dim o12 As Object
    Dim v13 As String
    Dim v14 As Byte
    Dim v15 As Byte
    Dim v16 As Boolean
    Dim v17 As Single
    Dim v18 As Single
    Dim v19 As Boolean
    Dim v20 As Byte
    Dim v21 As Boolean
    Set o1 = p1
    v2 = p2
    v3 = p3
    v4 = p4
    v5 = p5
    v6 = p6
    v7 = p7
    v8 = p8
    v9 = p9
    v10 = p10
    v11 = p11
    Set o12 = p12
    v13 = p13
    v14 = p14
    v15 = p15
    v16 = p16
    v17 = p17
    v18 = p18
    v19 = p19
    v20 = p20
    v21 = p21
    W_10595 = True
End Function

Public Function W_11274(p1 As Currency, p2 As Integer) As Single
    Dim v1 As Currency
    Dim v2 As Integer
    v1 = p1
    v2 = p2
    W_11274 = 1.5
End Function

Public Function W_11585(ByVal p1 As String, ByVal p2 As String, ByVal p3 As Integer, Optional ByVal p4 As Boolean = True, Optional ByVal p5 As String, Optional ByRef p6 As Application) As Integer
    Dim v1 As String
    Dim v2 As String
    Dim v3 As Integer
    Dim v4 As Boolean
    Dim v5 As String
    Dim o6 As Application
    v1 = p1
    v2 = p2
    v3 = p3
    v4 = p4
    v5 = p5
    Set o6 = p6
    W_11585 = 1234
End Function

Public Sub W_11703(ByRef p1 As Object, ByRef p2 As Variant, ByRef p3 As Variant, ByRef p4 As Collection, ByRef p5() As String, ByRef p6 As Long, ByRef p7 As Collection, ByRef p8 As Collection, ByRef p9 As Collection, ByRef p10 As Workbook)
    Dim o1 As Object
    Dim v2 As Variant
    Dim v3 As Variant
    Dim o4 As Collection
    Dim v6 As Long
    Dim o7 As Collection
    Dim o8 As Collection
    Dim o9 As Collection
    Dim o10 As Workbook
    Set o1 = p1
    v2 = p2
    v3 = p3
    Set o4 = p4
    v6 = p6
    Set o7 = p7
    Set o8 = p8
    Set o9 = p9
    Set o10 = p10
End Sub

Public Function W_16316(p1 As String, Optional p2 As Worksheet, Optional p3 As String = "WILD", Optional p4 As Long = &H11223344, Optional p5 As Long = &H11223344, Optional p6 As Long = &H11223344, Optional p7 As Variant, Optional p8 As Boolean = True, Optional p9 As Long = &H11223344, Optional p10 As String = "WILD", Optional ByRef p11 As Range) As Range
    Dim v1 As String
    Dim o2 As Worksheet
    Dim v3 As String
    Dim v4 As Long
    Dim v5 As Long
    Dim v6 As Long
    Dim v7 As Variant
    Dim v8 As Boolean
    Dim v9 As Long
    Dim v10 As String
    Dim o11 As Range
    v1 = p1
    Set o2 = p2
    v3 = p3
    v4 = p4
    v5 = p5
    v6 = p6
    v7 = p7
    v8 = p8
    v9 = p9
    v10 = p10
    Set o11 = p11
    Set W_16316 = Nothing
End Function

Public Function W_16317(Optional ByRef p1 As Worksheet, Optional p2 As String = "WILD") As Worksheet
    Dim o1 As Worksheet
    Dim v2 As String
    Set o1 = p1
    v2 = p2
    Set W_16317 = Nothing
End Function

Public Sub W_18883(ByRef p1 As Chart, ByRef p2 As Object)
    Dim o1 As Chart
    Dim o2 As Object
    Set o1 = p1
    Set o2 = p2
End Sub

Public Function W_1953(ByVal p1 As Workbook, ByVal p2 As String, ParamArray p3() As Variant) As Variant
    Dim o1 As Workbook
    Dim v2 As String
    Set o1 = p1
    v2 = p2
    W_1953 = 42
End Function

Public Function W_20062(ByVal p1 As Collection, p2 As Object, p3 As Collection, p4 As Collection, ByRef p5 As Long) As Collection
    Dim o1 As Collection
    Dim o2 As Object
    Dim o3 As Collection
    Dim o4 As Collection
    Dim v5 As Long
    Set o1 = p1
    Set o2 = p2
    Set o3 = p3
    Set o4 = p4
    v5 = p5
    Set W_20062 = Nothing
End Function

Public Function W_21984(ByRef p1() As Double, ByRef p2() As Date, ByVal p3 As Date, ByVal p4 As Long) As Double
    Dim v3 As Date
    Dim v4 As Long
    v3 = p3
    v4 = p4
    W_21984 = 2748.5
End Function

Public Function W_22379(p1 As LongLong, ByRef p2 As LongLong) As Boolean
    Dim v1 As LongLong
    Dim v2 As LongLong
    v1 = p1
    v2 = p2
    W_22379 = True
End Function

Public Function W_2331(ByVal p1 As String, Optional ByVal p2 As Boolean = True, Optional p3 As Application) As Object
    Dim v1 As String
    Dim v2 As Boolean
    Dim o3 As Application
    v1 = p1
    v2 = p2
    Set o3 = p3
    Set W_2331 = Nothing
End Function

Public Function W_25062(ByVal p1 As Application) As Workbook
    Dim o1 As Application
    Set o1 = p1
    Set W_25062 = Nothing
End Function

Public Function W_2636(p1 As Worksheet, p2 As String) As Chart
    Dim o1 As Worksheet
    Dim v2 As String
    Set o1 = p1
    v2 = p2
    Set W_2636 = Nothing
End Function

Public Function W_30878(ByVal p1 As Object, ByRef p2() As LongLong) As Object
    Dim o1 As Object
    Set o1 = p1
    Set W_30878 = Nothing
End Function

Public Function W_33088(ByRef p1() As LongPtr) As LongPtr
    W_33088 = 1234567890123^
End Function

Public Function W_3703(p1 As Date, p2 As Integer, p3 As Double, Optional p4 As Double, Optional p5 As Date, Optional p6 As Date, Optional p7 As Long) As Date
    Dim v1 As Date
    Dim v2 As Integer
    Dim v3 As Double
    Dim v4 As Double
    Dim v5 As Date
    Dim v6 As Date
    Dim v7 As Long
    v1 = p1
    v2 = p2
    v3 = p3
    v4 = p4
    v5 = p5
    v6 = p6
    v7 = p7
    W_3703 = #1/2/2020#
End Function

Public Function W_3832(ByVal p1 As Worksheet, ByVal p2 As String, ByVal p3 As Range, ByVal p4 As Variant, Optional ByRef p5 As Boolean = True) As Object
    Dim o1 As Worksheet
    Dim v2 As String
    Dim o3 As Range
    Dim v4 As Variant
    Dim v5 As Boolean
    Set o1 = p1
    v2 = p2
    Set o3 = p3
    v4 = p4
    v5 = p5
    Set W_3832 = Nothing
End Function

Public Sub W_3908(ByVal p1 As Chart, ByVal p2 As String, ByVal p3 As Range, ByVal p4 As Range, ByVal p5 As Long, ByVal p6 As Double, ByVal p7 As Boolean)
    Dim o1 As Chart
    Dim v2 As String
    Dim o3 As Range
    Dim o4 As Range
    Dim v5 As Long
    Dim v6 As Double
    Dim v7 As Boolean
    Set o1 = p1
    v2 = p2
    Set o3 = p3
    Set o4 = p4
    v5 = p5
    v6 = p6
    v7 = p7
End Sub

Public Sub W_4506(p1 As Long, p2() As Long, p3 As Long, p4() As Long, p5() As Long, p6 As Long, p7() As Boolean, p8() As Long, p9 As Long)
    Dim v1 As Long
    Dim v3 As Long
    Dim v6 As Long
    Dim v9 As Long
    v1 = p1
    v3 = p3
    v6 = p6
    v9 = p9
End Sub

Public Sub W_4632(p1 As Chart, Optional p2 As Boolean = True, Optional p3 As Variant)
    Dim o1 As Chart
    Dim v2 As Boolean
    Dim v3 As Variant
    Set o1 = p1
    v2 = p2
    v3 = p3
End Sub

Public Function W_5467(ByRef p1() As Currency, ByVal p2 As Long) As Object
    Dim v2 As Long
    v2 = p2
    Set W_5467 = Nothing
End Function

Public Function W_5780(p1 As Range, Optional p2 As Workbook, Optional p3 As Integer = 1234) As String
    Dim o1 As Range
    Dim o2 As Workbook
    Dim v3 As Integer
    Set o1 = p1
    Set o2 = p2
    v3 = p3
    W_5780 = "WILD"
End Function

Public Function W_5825(Optional p1 As Object, Optional p2 As Boolean) As Application
    Dim o1 As Object
    Dim v2 As Boolean
    Set o1 = p1
    v2 = p2
    Set W_5825 = Nothing
End Function

Public Sub W_6194(ByRef p1() As Byte, ByRef p2 As Long, ByRef p3 As Long, ByVal p4 As Integer, ByVal p5 As Integer, ByRef p6() As Integer, ByRef p7() As Object, ByRef p8() As Byte, ByVal p9 As Integer)
    Dim v2 As Long
    Dim v3 As Long
    Dim v4 As Integer
    Dim v5 As Integer
    Dim v9 As Integer
    v2 = p2
    v3 = p3
    v4 = p4
    v5 = p5
    v9 = p9
End Sub

Public Sub T_WildDrive()
    W_Init
    Call W_0401(Nothing, 1234567890123^, 1234567890123^, &H11223344)
    Call W_0442(gaSingle, &H11223344, &H11223344, &H11223344)
    Call W_0442(gaSingle)
    Call W_0448(2748.5)
    Call W_0951(9.99@)
    Call W_1023(1234567890123^)
    Call W_1027(1234567890123^, 9.99@)
    Call W_10595(Nothing, 42, &H11223344, &H11223344, &H11223344, &H11223344, &H11223344, &H11223344, &H11223344, &H11223344, &H11223344, Nothing, "WILD", 200, 200, True, 1.5, 1.5, True, 200, True)
    Call W_10595(Nothing, 42, &H11223344, &H11223344, &H11223344, &H11223344)
    Call W_11274(9.99@, 1234)
    Call W_11585("WILD", "WILD", 1234, True, "WILD", Nothing)
    Call W_11585("WILD", "WILD", 1234)
    Call W_11703(Nothing, 42, 42, Nothing, gaString, &H11223344, Nothing, Nothing, Nothing, Nothing)
    Call W_16316("WILD", Nothing, "WILD", &H11223344, &H11223344, &H11223344, 42, True, &H11223344, "WILD", Nothing)
    Call W_16316("WILD")
    Call W_16317(Nothing, "WILD")
    Call W_16317()
    Call W_18883(Nothing, Nothing)
    Call W_1953(Nothing, "WILD", 7, 8)
    Call W_1953(Nothing, "WILD")
    Call W_20062(Nothing, Nothing, Nothing, Nothing, &H11223344)
    Call W_21984(gaDouble, gaDate, #1/2/2020#, &H11223344)
    Call W_22379(1234567890123^, 1234567890123^)
    Call W_2331("WILD", True, Nothing)
    Call W_2331("WILD")
    Call W_25062(Nothing)
    Call W_2636(Nothing, "WILD")
    Call W_30878(Nothing, gaLL)
    Call W_33088(gaLL)
    Call W_3703(#1/2/2020#, 1234, 2748.5, 2748.5, #1/2/2020#, #1/2/2020#, &H11223344)
    Call W_3703(#1/2/2020#, 1234, 2748.5)
    Call W_3832(Nothing, "WILD", Nothing, 42, True)
    Call W_3832(Nothing, "WILD", Nothing, 42)
    Call W_3908(Nothing, "WILD", Nothing, Nothing, &H11223344, 2748.5, True)
    Call W_4506(&H11223344, gaLong, &H11223344, gaLong, gaLong, &H11223344, gaBool, gaLong, &H11223344)
    Call W_4632(Nothing, True, 42)
    Call W_4632(Nothing)
    Call W_5467(gaCur, &H11223344)
    Call W_5780(Nothing, Nothing, 1234)
    Call W_5780(Nothing)
    Call W_5825(Nothing, True)
    Call W_5825()
    Call W_6194(gaByte, &H11223344, &H11223344, 1234, 1234, gaInt, gaObj, gaByte, 1234)
End Sub
'@
     Invoke=@{ Name='T_WildDrive'; Args=@() }
     Expect={ param($t)
        if ($t.faults -gt 0) { return "$($t.faults) guarded reads faulted" }
        $e = @{}
        foreach ($r in $t.rows) {
            # 'W_*', not 'W_0*': the names are zero-padded to four digits, so W_1007 exists.
            if ($r.kind -eq 'entry' -and $r.source -eq 'VBA' -and $r.function -like 'W_*') {
                if (-not $e.ContainsKey($r.function)) { $e[$r.function] = @() }
                $e[$r.function] += $r }
        }
        # `a<slot>` is the trace's slot, not the parameter ordinal: a ByVal Variant takes three.
        # Only predictable values are listed, so no arrays and no typed Optional defaults.
        $want = @(
            @{ fn='W_0401'; args='a1=Nothing a2=1234567890123 a3=1234567890123 a4=287454020' }
            @{ fn='W_0442'; args='a2=287454020 a3=287454020 a4=287454020' }
            @{ fn='W_0448'; args='a1=2748.5' }
            @{ fn='W_0951'; args='a1=9.9900' }
            @{ fn='W_1023'; args='a1=1234567890123' }
            @{ fn='W_1027'; args='a1=1234567890123 a2=9.9900' }
            @{ fn='W_10595'; args='a1=Nothing a2=Integer(42) a5=287454020 a6=287454020 a7=287454020 a8=287454020 a9=287454020 a10=287454020 a11=287454020 a12=287454020 a13=287454020 a14=Nothing a15="WILD" a16=200 a17=200 a18=-1 a19=1.5 a20=1.5 a21=-1 a22=200 a23=-1' }
            @{ fn='W_10595'; args='a1=Nothing a2=Integer(42) a5=287454020 a6=287454020 a7=287454020 a8=287454020' }
            @{ fn='W_11274'; args='a1=9.9900 a2=1234' }
            @{ fn='W_11585'; args='a1="WILD" a2="WILD" a3=1234 a4=-1 a5="WILD"' }
            @{ fn='W_11585'; args='a1="WILD" a2="WILD" a3=1234' }
            @{ fn='W_11703'; args='a1=Nothing a2=Integer(42) a3=Integer(42) a6=287454020' }
            @{ fn='W_16316'; args='a1="WILD" a3="WILD" a4=287454020 a5=287454020 a6=287454020 a7=Integer(42) a8=-1 a9=287454020 a10="WILD"' }
            @{ fn='W_16316'; args='a1="WILD" a7=Missing' }
            @{ fn='W_16317'; args='a2="WILD"' }
            @{ fn='W_18883'; args='a2=Nothing' }
            @{ fn='W_1953'; args='a2="WILD" a3=Variant[0..1]{Integer(7),Integer(8)}' }
            @{ fn='W_1953'; args='a2="WILD"' }
            @{ fn='W_20062'; args='a2=Nothing a5=287454020' }
            @{ fn='W_21984'; args='a3=43832 a4=287454020' }
            @{ fn='W_22379'; args='a1=1234567890123 a2=1234567890123' }
            @{ fn='W_2331'; args='a1="WILD" a2=-1' }
            @{ fn='W_2331'; args='a1="WILD"' }
            @{ fn='W_2636'; args='a2="WILD"' }
            @{ fn='W_30878'; args='a1=Nothing' }
            @{ fn='W_3703'; args='a1=43832 a2=1234 a3=2748.5 a4=2748.5 a5=43832 a6=43832 a7=287454020' }
            @{ fn='W_3703'; args='a1=43832 a2=1234 a3=2748.5' }
            @{ fn='W_3832'; args='a2="WILD" a4=Integer(42) a7=-1' }
            @{ fn='W_3832'; args='a2="WILD" a4=Integer(42)' }
            @{ fn='W_3908'; args='a2="WILD" a5=287454020 a6=2748.5 a7=-1' }
            @{ fn='W_4506'; args='a1=287454020 a3=287454020 a6=287454020 a9=287454020' }
            @{ fn='W_4632'; args='a2=-1 a3=Integer(42)' }
            @{ fn='W_4632'; args='a3=Missing' }
            @{ fn='W_5467'; args='a2=287454020' }
            @{ fn='W_5780'; args='a3=1234' }
            @{ fn='W_5825'; args='a1=Nothing a2=-1' }
            @{ fn='W_6194'; args='a2=287454020 a3=287454020 a4=1234 a5=1234 a9=1234' }
        )
        # matched to any satisfying row, so the order Excel wrote them in cannot fail the test
        $pool = @{}
        foreach ($k in $e.Keys) { $pool[$k] = [System.Collections.ArrayList]::new($e[$k]) }
        foreach ($w in $want) {
            if (-not $pool.ContainsKey($w.fn)) { return "$($w.fn) never entered" }
            $hit = -1
            for ($i = 0; $i -lt $pool[$w.fn].Count; $i++) {
                $got = [string]$pool[$w.fn][$i].args
                $ok = $true
                foreach ($tok in $w.args.Split(' ')) {
                    if (-not $tok) { continue }
                    $n2 = $tok.Split('=')[0]
                    $v2 = $tok.Substring($tok.IndexOf('=') + 1)
                    # read to the next key: the type sits between, and a value can hold anything
                    $rx = '(?:^|\s)' + [regex]::Escape($n2) + '(?::[^=\s]*)?=(.*?)(?=\s+a\d+(?::[^=\s]*)?=|$)'
                    $m2 = [regex]::Match($got, $rx)
                    if (-not $m2.Success -or $m2.Groups[1].Value -ne $v2) { $ok = $false; break }
                }
                if ($ok) { $hit = $i; break }
            }
            if ($hit -lt 0) {
                return "$($w.fn): no entry row matched planted [$($w.args)]; first row was [$([string]$pool[$w.fn][0].args)]" }
            $pool[$w.fn].RemoveAt($hit)
        }
        if ($t.framesOpened -ne $t.framesClosed) {
            return "LEAK: opened $($t.framesOpened), closed $($t.framesClosed)" }
        $null }
     Why='parameter lists nobody here invented -- the shapes real VBA actually
          declares -- with planted values read back out of the argument walker' }

if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-VbaCase $case
