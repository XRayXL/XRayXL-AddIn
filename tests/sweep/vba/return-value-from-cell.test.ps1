# A VBA exit row carries its activation's result, for every declared type, array and Variant,
# at any depth. Cell-visible values are checked against what Excel put in the cell, deeper ones
# against planted constants.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    $srcM = @'
Private R_ObjRoundTripCount As Long

' ---- SCALARS, one per exit slot, each planted so no other decode fits ----
Public Function R_Double() As Double
    R_Double = 1234.5
End Function
Public Function R_Locals() As Double
    Dim a As Double, b As Double
    a = 8877.25
    b = 999.5
    R_Locals = 4321.75
End Function
Public Function R_Args(ByVal p As Double, ByVal q As Double) As Double
    R_Args = p * q
End Function
Public Function R_Negative(ByVal p As Double) As Double
    R_Negative = -(p / 4)
End Function
Public Function R_Single() As Single
    R_Single = 1234.5
End Function
Public Function R_Byte() As Byte
    R_Byte = 200
End Function
Public Function R_Integer() As Integer
    R_Integer = 4321
End Function
Public Function R_IntegerNeg() As Integer
    R_IntegerNeg = -4321
End Function
Public Function R_Long() As Long
    R_Long = 123456789
End Function
Public Function R_LongNeg() As Long
    R_LongNeg = -123456789
End Function
Public Function R_LongLong() As LongLong
    R_LongLong = 1234567890123
End Function
Public Function R_Currency() As Currency
    R_Currency = 12.3456
End Function
Public Function R_Date() As Date
    R_Date = #1/2/2020#
End Function
Public Function R_Bool() As Boolean
    R_Bool = True
End Function
Public Function R_BoolFalse() As Boolean
    R_BoolFalse = False
End Function

' ---- VARIANTS: the live VARIANT at [R14-0x18], decoded by its vt ----
Public Function R_Variant() As Variant
    R_Variant = 1234.5
End Function
Public Function R_VariantStr() As Variant
    R_VariantStr = "VARSTR"
End Function
Public Function R_VariantLong() As Variant
    R_VariantLong = CLng(777)
End Function
Public Function R_VariantEmpty() As Variant
    Dim a As Double
    a = 1
End Function
' Decimal exists only inside a Variant, via CDec. 28 significant digits: more
' than a Double holds, so the trace must carry them exactly while the cell
' shows an approximation.
Public Function R_VariantDec() As Variant
    ' digits only, then an exact division: a decimal point in the string would follow the system locale
    R_VariantDec = CDec("12345678901234567890123456") / CDec("1000000000000000000000")
End Function
Public Function R_VariantDecNeg() As Variant
    R_VariantDecNeg = CDec(-42)
End Function
Public Function R_ArrDec() As Variant
    R_ArrDec = Array(CDec(3) / CDec(2), CDec(-9) / CDec(4))
End Function
Public Function R_ArrVar() As Variant
    Dim a(0 To 2) As Double
    a(0) = 1234.5
    a(1) = 2
    a(2) = 3
    R_ArrVar = a
End Function

' ---- TYPED ARRAYS: a live SAFEARRAY* at [R14-8], self-typed ----
Public Function R_ArrDbl() As Double()
    Dim a(0 To 2) As Double
    a(0) = 1234.5
    a(1) = 2
    a(2) = 3
    R_ArrDbl = a
End Function
Public Function R_ArrLng() As Long()
    Dim a(1 To 3) As Long
    a(1) = 1234
    a(2) = 5
    a(3) = 6
    R_ArrLng = a
End Function
Public Function R_ArrStr() As String()
    Dim a(0 To 1) As String
    a(0) = "XRAYRET"
    a(1) = "TWO"
    R_ArrStr = a
End Function
' 2-D: VBA stores column-major, but the trace reads row by row: a(0,0), a(0,1), a(1,0), a(1,1)
Public Function R_Arr2D() As Double()
    Dim a(0 To 1, 0 To 1) As Double
    a(0, 0) = 1234.5
    a(0, 1) = 2
    a(1, 0) = 3
    a(1, 1) = 4
    R_Arr2D = a
End Function
' Larger than the four elements the trace shows, so the "..." path is exercised
Public Function R_ArrBig() As Long()
    Dim a(0 To 9) As Long
    Dim i As Long
    For i = 0 To 9
        a(i) = 100 + i
    Next i
    R_ArrBig = a
End Function

' ---- THE STACK: results at depth 2 and 3, and a Sub at depth 2 ----
Public Function R_Outer() As Double
    R_Outer = R_Mid() + 1
End Function
Private Function R_Mid() As Double
    R_Mid = R_Inner() + 10
End Function
Private Function R_Inner() As Double
    R_Inner = 100.25
End Function
Public Function R_OuterL() As Long
    R_OuterL = R_InnerL() * 2
End Function
Private Function R_InnerL() As Long
    R_InnerL = 21
End Function
' A Sub at depth 2 whose first local is a normal double -- the case the old
' caller gate existed for. Its exit slot (635) says "no result".
Public Function R_CallsSub() As Double
    S_Local
    R_CallsSub = 7
End Function
Private Sub S_Local()
    Dim a As Double
    a = 5551.25
End Sub
' A Sub at the TOP, reached by Application.Run
Public Sub R_SubTop()
    Dim a As Double
    a = 6661.5
End Sub

' ---- VARIANT ARRAYS OF VARIANTS: what Array() and Range.Value produce ----
' 24-byte elements, each a VARIANT of its own -- so a nested array or an
' object inside is just an element whose own vt says so.
Public Function R_ArrMixed() As Variant
    R_ArrMixed = Array(1234.5, "two", CLng(3), True)
End Function
Public Function R_ArrNested() As Variant
    R_ArrNested = Array(Array(1234.5, 2), 3)
End Function
' Range.Value: a 2-D Variant array read from constant cells R1:S2.
Public Function R_RangeVal() As Variant
    R_RangeVal = ThisWorkbook.Worksheets(1).Range("R1:S2").Value
End Function

' ---- OBJECTS: as a declared return, in a Variant, in an array, and as an
' ---- argument handed down and handed back -- the same address on both rows.
Public Function R_Obj() As Object
    Set R_Obj = New Collection
End Function
Public Function R_ObjNothing() As Object
    Set R_ObjNothing = Nothing
End Function
Public Function R_VarObj() As Variant
    Set R_VarObj = New Collection
End Function
Public Function R_ArrWithObj() As Variant
    Dim c As New Collection
    R_ArrWithObj = Array(c, 1234.5)
End Function
Public Function R_ObjRoundTrip() As Double
    Dim c As New Collection
    c.Add 7
    R_ObjRoundTrip = R_TakesObj(c)
End Function
Private Function R_TakesObj(ByVal o As Collection) As Object
    Set R_TakesObj = o
    R_ObjRoundTripCount = o.Count
End Function

' ---- STRINGS: a live BSTR at [R14-8], at any depth, and the empty one ----
Public Function R_Str() As String
    R_Str = "XRAYRET"
End Function
Public Function R_StrEmpty() As String
    R_StrEmpty = ""
End Function
Public Function R_StrOuter() As String
    R_StrOuter = R_StrInner() & "-outer"
End Function
Private Function R_StrInner() As String
    R_StrInner = "inner"
End Function
'@
    # array returns spill, so they sit one per column, well apart
    New-XRayMacroBook $sx 'VbaRet' @(
        @{ Kind=1; Name='M'; Code=$srcM }
    ) @{} {
        param($ws)
        $cells = [ordered]@{
            'A1'  = '=R_Double()';      'A2'  = '=R_Locals()';     'A3'  = '=R_Args(3,4.5)'
            'A4'  = '=R_Negative(9)';   'A5'  = '=R_Single()';     'A6'  = '=R_Byte()'
            'A7'  = '=R_Integer()';     'A8'  = '=R_IntegerNeg()'; 'A9'  = '=R_Long()'
            'A10' = '=R_LongNeg()';     'A11' = '=R_LongLong()';   'A12' = '=R_Currency()'
            'A13' = '=R_Date()';        'A14' = '=R_Bool()';       'A15' = '=R_BoolFalse()'
            'A16' = '=R_Variant()';     'A17' = '=R_VariantStr()'; 'A18' = '=R_VariantLong()'
            'A19' = '=R_VariantEmpty()';'A20' = '=R_Outer()';      'A21' = '=R_OuterL()'
            'A22' = '=R_CallsSub()';    'A23' = '=R_Str()';       'A24' = '=R_StrEmpty()'
            'A25' = '=R_StrOuter()';   'A26' = '=R_ObjRoundTrip()'
            'A27' = '=R_VariantDec()'; 'A28' = '=R_VariantDecNeg()'
            'R1' = '10'; 'R2' = '20'; 'S1' = 'x'; 'S2' = 'TRUE'
            'AE1' = '=R_Obj()'; 'AG1' = '=R_ObjNothing()'; 'AI1' = '=R_VarObj()'; 'AK1' = '=R_ArrWithObj()'
            'C1' = '=R_ArrVar()'; 'E1' = '=R_ArrDbl()'; 'G1' = '=R_ArrLng()'
            'I1' = '=R_ArrStr()'; 'K1' = '=R_Arr2D()';  'M1' = '=R_ArrBig()'
            'O1' = '=R_ArrMixed()'; 'Q1' = '=R_ArrNested()'; 'U1' = '=R_RangeVal()'
            'W1' = '=R_ArrDec()'
        }
        foreach ($k in $cells.Keys) { $ws.Range($k).Formula = $cells[$k] }
        $ws.Range('S2').Value2 = $true
    }
    $book = Get-XRayMacroBook
    $ws = $book.Sheet; $bookPath = $book.Path

    $mark = Get-LogLength $paths.Log
    $pressed = Invoke-XRayCommand $sx 'XRayXL_Arm'
    if ($pressed -ne 'pressed') { Complete-Test -Fail -Detail "arm: $pressed" }
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    Invoke-XRayRecalc $app 'Rebuild'
    $app.Run(("'{0}'!R_SubTop" -f (Split-Path $bookPath -Leaf)))
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $rows = Select-BookRows (Read-TraceRows $sx.ProcId) (Split-Path $bookPath -Leaf)

    function Get-Return([string]$Fn) {
        $e = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq $Fn }) | Select-Object -First 1
        if (-not $e) { return $null }
        return @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') -and $_.span -eq $e.span }) | Select-Object -First 1
    }
    # ---- scalars, against what Excel computed --------------------------
    foreach ($c in @(
        @{ Fn='R_Double';     Cell='A1';  Type='Double'   }
        @{ Fn='R_Locals';     Cell='A2';  Type='Double'   }
        @{ Fn='R_Args';       Cell='A3';  Type='Double'   }
        @{ Fn='R_Negative';   Cell='A4';  Type='Double'   }
        @{ Fn='R_Single';     Cell='A5';  Type='Single'   }
        @{ Fn='R_Byte';       Cell='A6';  Type='Byte'     }
        @{ Fn='R_Integer';    Cell='A7';  Type='Integer'  }
        @{ Fn='R_IntegerNeg'; Cell='A8';  Type='Integer'  }
        @{ Fn='R_Long';       Cell='A9';  Type='Long'     }
        @{ Fn='R_LongNeg';    Cell='A10'; Type='Long'     }
        @{ Fn='R_LongLong';   Cell='A11'; Type='LongLong' }
        @{ Fn='R_Currency';   Cell='A12'; Type='Currency' }
        @{ Fn='R_Date';       Cell='A13'; Type='Double'   }
        @{ Fn='R_Outer';      Cell='A20'; Type='Double'   }   # depth 1 of a 3-deep chain
        @{ Fn='R_OuterL';     Cell='A21'; Type='Long'     }
        @{ Fn='R_CallsSub';   Cell='A22'; Type='Double'   }
    )) {
        $x = Get-Return $c.Fn
        if (-not $x) { Check "$($c.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        $want = [double]$ws.Range($c.Cell).Value2
        if (-not $x.ret) { Check "$($c.Fn)-returns-value" $false "ret is empty; the cell says $want"; continue }
        $tol = if ($c.Type -eq 'Single') { [math]::Max([math]::Abs($want) * 1e-6, 1e-6) } else { 1e-9 }
        $ok = ([math]::Abs(([double]$x.ret) - $want) -le $tol) -and ($x.rettype -eq $c.Type)
        Check "$($c.Fn)-returns-value" $ok ("trace='{0}' rettype='{1}' cell={2}" -f $x.ret, $x.rettype, $want)
    }

    # ---- booleans: I2, so -1/0 --------------------------------------------
    foreach ($b in @(@{ Fn='R_Bool'; Want='-1' }, @{ Fn='R_BoolFalse'; Want='0' })) {
        $x = Get-Return $b.Fn
        if (-not $x) { Check "$($b.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        Check "$($b.Fn)-reports-i2" (($x.ret -eq $b.Want) -and ($x.rettype -eq 'Integer')) ("ret='{0}' rettype='{1}'" -f $x.ret, $x.rettype)
    }

    # ---- the stack: depth 2 and 3, against their planted constants -------
    foreach ($d in @(
        @{ Fn='R_Inner';  Want='100.25'; Type='Double' }   # depth 3
        @{ Fn='R_Mid';    Want='110.25'; Type='Double' }   # depth 2
        @{ Fn='R_InnerL'; Want='21';     Type='Long'   }   # depth 2, a Long
    )) {
        $x = Get-Return $d.Fn
        if (-not $x) { Check "$($d.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        $ok = ($x.ret -and ([math]::Abs(([double]$x.ret) - [double]$d.Want) -lt 1e-9) -and ($x.rettype -eq $d.Type))
        Check "$($d.Fn)-returns-at-depth" $ok ("ret='{0}' rettype='{1}' expected {2}" -f $x.ret, $x.rettype, $d.Want)
    }

    # ---- Variants: the held value, a number other than Double named ------------
    foreach ($v in @(
        @{ Fn='R_Variant';      Want='1234.5'      }
        @{ Fn='R_VariantStr';   Want='"VARSTR"'    }
        @{ Fn='R_VariantLong';  Want='Long(777)'   }
        @{ Fn='R_VariantEmpty'; Want='Empty'       }
        @{ Fn='R_VariantDec';   Want='Decimal(12345.678901234567890123456)' }
        @{ Fn='R_VariantDecNeg';Want='Decimal(-42)' }
        @{ Fn='R_ArrVar';       Want='Double[0..2]{1234.5,2,3}' }
    )) {
        $x = Get-Return $v.Fn
        if (-not $x) { Check "$($v.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        Check "$($v.Fn)-variant-decoded" (($x.ret -eq $v.Want) -and ($x.rettype -eq 'Variant')) ("ret='{0}' rettype='{1}' expected {2}" -f $x.ret, $x.rettype, $v.Want)
    }

    # ---- Variant arrays of Variants, nested arrays, Range.Value ---------------
    foreach ($v in @(
        @{ Fn='R_ArrMixed';  Want='Variant[0..3]{1234.5,"two",Long(3),TRUE}'   }
        @{ Fn='R_ArrNested'; Want='Variant[0..1]{Variant[0..1]{1234.5,Integer(2)},Integer(3)}' }
        @{ Fn='R_ArrDec';    Want='Variant[0..1]{Decimal(1.5),Decimal(-2.25)}' }
        @{ Fn='R_RangeVal';  Want='Variant[1..2,1..2]{{10,"x"},{20,TRUE}}' }   # a level per row: {R1,S1},{R2,S2}
    )) {
        $x = Get-Return $v.Fn
        if (-not $x) { Check "$($v.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        Check "$($v.Fn)-variant-array-decoded" (($x.ret -eq $v.Want) -and ($x.rettype -eq 'Variant')) ("ret='{0}' rettype='{1}' expected {2}" -f $x.ret, $x.rettype, $v.Want)
    }

    # the address is what follows one object from an argument to a result
    $emptyColl = '^Collection@0x[0-9A-Fa-f]+=Variant\[1\.\.0\]\{\}$'
    $x = Get-Return 'R_Obj'
    if (-not $x) { Check 'R_Obj-has-exit-row' $false 'no paired exit row' }
    else { Check 'R_Obj-object-returned' (($x.ret -match $emptyColl) -and ($x.rettype -eq 'Object')) ("ret='{0}' rettype='{1}'" -f $x.ret, $x.rettype) }
    $x = Get-Return 'R_ObjNothing'
    if (-not $x) { Check 'R_ObjNothing-has-exit-row' $false 'no paired exit row' }
    else { Check 'R_ObjNothing-is-Nothing' (($x.ret -eq 'Nothing') -and ($x.rettype -eq 'Object')) ("ret='{0}' rettype='{1}'" -f $x.ret, $x.rettype) }
    $x = Get-Return 'R_VarObj'
    if (-not $x) { Check 'R_VarObj-has-exit-row' $false 'no paired exit row' }
    else { Check 'R_VarObj-object-in-variant' (($x.ret -match $emptyColl) -and ($x.rettype -eq 'Variant')) ("ret='{0}' rettype='{1}'" -f $x.ret, $x.rettype) }
    $x = Get-Return 'R_ArrWithObj'
    if (-not $x) { Check 'R_ArrWithObj-has-exit-row' $false 'no paired exit row' }
    else { Check 'R_ArrWithObj-object-in-array' (($x.ret -match '^Variant\[0\.\.1\]\{Collection@0x[0-9A-Fa-f]+=Variant\[1\.\.0\]\{\},1234\.5\}$') -and ($x.rettype -eq 'Variant')) ("ret='{0}'" -f $x.ret) }

    # the object handed down as an argument comes back as the result, at one address
    $e = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'R_TakesObj' }) | Select-Object -First 1
    $x = Get-Return 'R_TakesObj'
    if (-not $e -or -not $x) { Check 'R_TakesObj-traced' $false 'entry or exit missing' }
    else {
        $argAddr = [regex]::Match($e.args, '@0x([0-9A-Fa-f]+)').Groups[1].Value
        $retAddr = [regex]::Match($x.ret,  '@0x([0-9A-Fa-f]+)').Groups[1].Value
        Check 'object-argument-and-return-are-the-same-address' ($argAddr -and ($argAddr -eq $retAddr) -and ($x.rettype -eq 'Object')) ("arg='{0}' ret='{1}'" -f $e.args, $x.ret)
    }

    # ---- typed arrays: shape, element type, and the elements ----------------
    foreach ($a in @(
        @{ Fn='R_ArrDbl'; Want='Double[0..2]{1234.5,2,3}';         Type='Double()' }
        @{ Fn='R_ArrLng'; Want='Long[1..3]{1234,5,6}';             Type='Long()'   }
        @{ Fn='R_ArrStr'; Want='String[0..1]{"XRAYRET","TWO"}';   Type='String()' }
        @{ Fn='R_Arr2D';  Want='Double[0..1,0..1]{{1234.5,2},{3,4}}'; Type='Double()' }   # a level per row
        @{ Fn='R_ArrBig'; Want='Long[0..9]{100,101,102,103,104,105,106,107,108,109}'; Type='Long()'   }
    )) {
        $x = Get-Return $a.Fn
        if (-not $x) { Check "$($a.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        Check "$($a.Fn)-array-decoded" (($x.ret -eq $a.Want) -and ($x.rettype -eq $a.Type)) ("ret='{0}' rettype='{1}' expected {2}" -f $x.ret, $x.rettype, $a.Want)
    }

    # a spilled array's first element is what the cell shows
    foreach ($s in @(@{ Fn='R_ArrDbl'; Cell='E1' }, @{ Fn='R_ArrLng'; Cell='G1' }, @{ Fn='R_Arr2D'; Cell='K1' })) {
        $x = Get-Return $s.Fn
        $first = Get-XRayCellText $ws.Range($s.Cell)
        $ok = ($x -and $x.ret -match ('\{' + [regex]::Escape($first) + '[,}]'))
        Check "$($s.Fn)-first-element-matches-cell" $ok ("cell={0} trace='{1}'" -f $first, $(if ($x) { $x.ret } else { '' }))
    }

    # ---- strings, against the cells --------------------------------------------
    foreach ($t in @(
        @{ Fn='R_Str';      Cell='A23'; Want='"XRAYRET"'    }
        @{ Fn='R_StrEmpty'; Cell='A24'; Want='""'            }
        @{ Fn='R_StrOuter'; Cell='A25'; Want='"inner-outer"' }
    )) {
        $x = Get-Return $t.Fn
        if (-not $x) { Check "$($t.Fn)-has-exit-row" $false 'no paired exit row'; continue }
        $cell = Get-XRayCellText $ws.Range($t.Cell)
        $ok = ($x.ret -eq $t.Want) -and ($x.rettype -eq 'String') -and ($x.ret -eq ('"' + $cell + '"'))
        Check "$($t.Fn)-string-decoded" $ok ("ret='{0}' rettype='{1}' cell='{2}'" -f $x.ret, $x.rettype, $cell)
    }
    # ...and a String returned at depth 2, against its planted value.
    $x = Get-Return 'R_StrInner'
    if (-not $x) { Check 'R_StrInner-has-exit-row' $false 'no paired exit row' }
    else { Check 'R_StrInner-string-at-depth' (($x.ret -eq '"inner"') -and ($x.rettype -eq 'String')) ("ret='{0}' rettype='{1}'" -f $x.ret, $x.rettype) }

    # ---- Subs ----------------------------------------------------------------

    foreach ($sub in @('S_Local', 'R_SubTop')) {
        $x = Get-Return $sub
        if (-not $x) { Check "$sub-has-exit-row" $false 'the Sub was not traced'; continue }
        Check "$sub-no-return" ((-not $x.ret) -and (-not $x.rettype)) ("ret='{0}' rettype='{1}' (a Sub has no result; its exit slot says so)" -f $x.ret, $x.rettype)
    }

    $earlyRet = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.ret })
    Check 'entry-rows-carry-no-return' ($earlyRet.Count -eq 0) "$($earlyRet.Count) entry row(s) with a ret"

    $p = @(Test-RowInvariants $rows)
    if ($p.Count) { Complete-Test -Fail -Detail ($p -join '; ') }

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails return case(s) failed" }
    Complete-Test -Pass -Detail 'scalars, booleans, strings, Variants, Variant arrays of Variants, nested arrays, objects in every position, depth 2 and 3 and Subs all as measured'
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
