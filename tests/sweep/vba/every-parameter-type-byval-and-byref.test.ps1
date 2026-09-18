# Every declared parameter type, ByVal and ByRef: twelve types, twenty-four cases, so the whole
# grid is asserted rather than the entries some procedure happened to reach.
#
# ByVal and ByRef of the same declared type reach the decoder differently (one slot holds the
# value, the other a pointer, through opcodes from different families) but describe the same
# planted value, so they must render the same text.
#
# The signature is asserted separately and loosely: the declared type must appear in it, with
# `&` for ByRef. That is what says the p-code walk recovered the type, since a BSTR and a
# SAFEARRAY prove themselves whatever the opcode table knows.
#
# `Date` is a `Double` and `Boolean` is an `Integer`: one opcode each, asserted as such.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$classCode = @'
Public Tag As Long
'@

$enumAndUdt = @'
Public Enum XRColour
    XRRed = 1
    XRBlue = 2
End Enum

Public Type XRPoint
    X As Long
    Y As Long
End Type
'@

$moduleCode = @'
' Each body READS its parameter, which is what makes the p-code emit a typed
' load. A parameter never read has no recoverable type -- that is a separate
' shape, covered separately.
Public Sub VByte(ByVal p As Byte)
    Dim t As Byte
    t = p
End Sub
Public Sub RByte(ByRef p As Byte)
    Dim t As Byte
    t = p
End Sub

Public Sub VInteger(ByVal p As Integer)
    Dim t As Integer
    t = p
End Sub
Public Sub RInteger(ByRef p As Integer)
    Dim t As Integer
    t = p
End Sub

Public Sub VLong(ByVal p As Long)
    Dim t As Long
    t = p
End Sub
Public Sub RLong(ByRef p As Long)
    Dim t As Long
    t = p
End Sub

Public Sub VLongLong(ByVal p As LongLong)
    Dim t As LongLong
    t = p
End Sub
Public Sub RLongLong(ByRef p As LongLong)
    Dim t As LongLong
    t = p
End Sub

Public Sub VSingle(ByVal p As Single)
    Dim t As Single
    t = p
End Sub
Public Sub RSingle(ByRef p As Single)
    Dim t As Single
    t = p
End Sub

Public Sub VDouble(ByVal p As Double)
    Dim t As Double
    t = p
End Sub
Public Sub RDouble(ByRef p As Double)
    Dim t As Double
    t = p
End Sub

Public Sub VCurrency(ByVal p As Currency)
    Dim t As Currency
    t = p
End Sub
Public Sub RCurrency(ByRef p As Currency)
    Dim t As Currency
    t = p
End Sub

Public Sub VString(ByVal p As String)
    Dim t As String
    t = p
End Sub
Public Sub RString(ByRef p As String)
    Dim t As String
    t = p
End Sub

Public Sub VObject(ByVal p As Object)
    Dim t As Object
    Set t = p
End Sub
Public Sub RObject(ByRef p As Object)
    Dim t As Object
    Set t = p
End Sub

Public Sub VVariant(ByVal p As Variant)
    Dim t As Variant
    t = p
End Sub
Public Sub RVariant(ByRef p As Variant)
    Dim t As Variant
    t = p
End Sub

Public Sub VDate(ByVal p As Date)
    Dim t As Date
    t = p
End Sub
Public Sub RDate(ByRef p As Date)
    Dim t As Date
    t = p
End Sub

Public Sub VBoolean(ByVal p As Boolean)
    Dim t As Boolean
    t = p
End Sub
Public Sub RBoolean(ByRef p As Boolean)
    Dim t As Boolean
    t = p
End Sub

' ---- LongPtr: an ALIAS, not a type of its own. On a 64-bit host it IS
' LongLong (MS-VBAL 2.1), so it must read as one -- and that is worth pinning,
' because an alias resolved differently would be invisible in the source.
Public Sub VLongPtr(ByVal p As LongPtr)
    Dim t As LongPtr
    t = p
End Sub
Public Sub RLongPtr(ByRef p As LongPtr)
    Dim t As LongPtr
    t = p
End Sub

' ---- Enum: underlying type Long.
Public Sub VEnum(ByVal p As XRColour)
    Dim t As XRColour
    t = p
End Sub
Public Sub REnum(ByRef p As XRColour)
    Dim t As XRColour
    t = p
End Sub

' ---- a CLASS MODULE type, not the generic Object.
Public Sub VClass(ByVal p As XRPType)
    Dim t As Long
    t = p.Tag
End Sub
Public Sub RClass(ByRef p As XRPType)
    Dim t As Long
    t = p.Tag
End Sub

' ---- an imported COM class.
Public Sub VCom(ByVal p As Collection)
    Dim t As Long
    t = p.Count
End Sub
Public Sub RCom(ByRef p As Collection)
    Dim t As Long
    t = p.Count
End Sub

' ---- a UDT. ByRef only: a UDT parameter is always by reference.
Public Sub RUdt(ByRef p As XRPoint)
    Dim t As Long
    t = p.X
End Sub

' ---- an ARRAY. ByRef only: `ByVal` on an array is a compile error
' (MS-VBAL 5.3.1.5), so there is no ByVal case to write.
Public Sub RArray(ByRef p() As Long)
    Dim t As Long
    t = p(1)
End Sub

' ---- ParamArray: always a resizable ByRef array of Variant.
Public Sub RParamArray(ParamArray p() As Variant)
    Dim t As Long
    t = UBound(p)
End Sub

' ---- A TYPED VARIABLE INTO A VARIANT PARAMETER. ByVal copies it; ByRef hands
' over a VARIANT tagged VT_BYREF that points at the caller's variable, so the
' callee's writes reach it. Both must read the variable's value.
Public Sub VVarFromStr(ByVal p As Variant)
    Dim t As Variant
    t = p
End Sub
Public Sub RVarFromStr(ByRef p As Variant)
    Dim t As Variant
    t = p
End Sub
Public Sub RVarFromLng(ByRef p As Variant)
    Dim t As Variant
    t = p
End Sub

' ---- DECIMAL exists only inside a Variant (CDec); it cannot be declared.
Public Sub VDec(ByVal p As Variant)
    Dim t As Variant
    t = p
End Sub
Public Sub RDec(ByRef p As Variant)
    Dim t As Variant
    t = p
End Sub

' ---- Optional, supplied and omitted.
Public Sub VOptional(Optional ByVal p As Long = 99)
    Dim t As Long
    t = p
End Sub

Public Sub Drive()
    Dim byt As Byte:      byt = 7
    Dim int16 As Integer: int16 = 1234
    Dim lng As Long:      lng = 123456
    Dim ll As LongLong:   ll = 4294967296#
    Dim sng As Single:    sng = 1.5
    Dim dbl As Double:    dbl = 2748.5
    Dim cur As Currency:  cur = 9.99
    Dim str1 As String:   str1 = "hello"
    Dim obj As Object:    Set obj = New XRPType
    Dim vnt As Variant:   vnt = 42
    Dim dt As Date:       dt = DateSerial(2020, 1, 1)
    Dim bl As Boolean:    bl = True
    Dim lp As LongPtr:    lp = 4294967296#
    Dim en As XRColour:   en = XRBlue
    Dim cls As XRPType:   Set cls = New XRPType: cls.Tag = 77
    Dim com As Collection: Set com = New Collection: com.Add 1
    Dim udt As XRPoint:   udt.X = 55: udt.Y = 66
    Dim arr(1 To 3) As Long: arr(1) = 111: arr(2) = 222: arr(3) = 333
    Dim dec As Variant:   dec = CDec(3) / CDec(2)   ' CDec("1.5") would read the system decimal separator

    VByte byt:        RByte byt
    VInteger int16:   RInteger int16
    VLong lng:        RLong lng
    VLongLong ll:     RLongLong ll
    VSingle sng:      RSingle sng
    VDouble dbl:      RDouble dbl
    VCurrency cur:    RCurrency cur
    VString str1:     RString str1
    VObject obj:      RObject obj
    VVariant vnt:     RVariant vnt
    VDate dt:         RDate dt
    VBoolean bl:      RBoolean bl
    VLongPtr lp:      RLongPtr lp
    VEnum en:         REnum en
    VClass cls:       RClass cls
    VCom com:         RCom com
    RUdt udt
    RArray arr
    RParamArray 1, 2, 3
    VOptional 5
    VOptional
    VVarFromStr str1: RVarFromStr str1: RVarFromLng lng
    VDec dec:         RDec dec
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    # Enum and Type are module-level declarations and must live in the
    # declarations section of a standard module, ahead of any procedure.
    New-XRayMacroBook $sx 'PType' @(
        @{ Kind=1; Name='PTypeCase'; Code=$moduleCode }
        @{ Kind=2; Name='XRPType'; Code=$classCode }
        @{ Kind=1; Name='XRPTypes'; Code=$enumAndUdt }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $rows = @(Read-TraceRows $sx.ProcId)
    function RowOf([string]$fn) {
        $r = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq $fn })
        if ($r.Count) { return $r[0] } else { return $null }
    }
    # The value with the `aN:Type=` label stripped, so ByVal and ByRef compare.
    function ValOf([string]$fn) {
        $r = RowOf $fn
        if (-not $r) { return '(no row)' }
        return ([string]$r.args) -replace '^a1(:[^=]*)?=', ''
    }

    # Declared type -> the name expected in the signature, and the planted value.
    # `Date` shares the Double opcode and `Boolean` shares Integer's, so those
    # two expect the shared name deliberately.
    $grid = @(
        @{ T='Byte';     V='VByte';     R='RByte';     Sig='Byte';     Want='7'          }
        @{ T='Integer';  V='VInteger';  R='RInteger';  Sig='Integer';  Want='1234'       }
        @{ T='Long';     V='VLong';     R='RLong';     Sig='Long';     Want='123456'     }
        @{ T='LongLong'; V='VLongLong'; R='RLongLong'; Sig='';         Want='4294967296' }
        @{ T='Single';   V='VSingle';   R='RSingle';   Sig='Single';   Want='1.5'        }
        @{ T='Double';   V='VDouble';   R='RDouble';   Sig='Double';   Want='2748.5'     }
        @{ T='Currency'; V='VCurrency'; R='RCurrency'; Sig='Currency'; Want='9.9900'     }
        @{ T='String';   V='VString';   R='RString';   Sig='String';   Want='"hello"'    }
        @{ T='Object';   V='VObject';   R='RObject';   Sig='Object';   Want=''           }
        @{ T='Variant';  V='VVariant';  R='RVariant';  Sig='Variant';  Want=''           }
        @{ T='Date';     V='VDate';     R='RDate';     Sig='Double';   Want=''           }
        @{ T='Boolean';  V='VBoolean';  R='RBoolean';  Sig='Integer';  Want='-1'         }
        # LongPtr is an ALIAS for LongLong on a 64-bit host, so it reads as one.
        @{ T='LongPtr';  V='VLongPtr';  R='RLongPtr';  Sig='';         Want='4294967296' }
        # An Enum's underlying type is Long.
        @{ T='Enum';     V='VEnum';     R='REnum';     Sig='Long';     Want='2'          }
        @{ T='Class';    V='VClass';    R='RClass';    Sig='Object';   Want=''           }
        @{ T='COM';      V='VCom';      R='RCom';      Sig='Object';   Want=''           }
    )

    Write-Output ''
    Write-Output ('  {0,-9} {1,-14} {2,-22} {3,-15} {4}' -f 'declared', 'ByVal sig', 'ByVal value', 'ByRef sig', 'ByRef value')
    foreach ($g in $grid) {
        Write-Output ('  {0,-9} {1,-14} {2,-22} {3,-15} {4}' -f `
            $g.T, (SigOf $rows $g.V), (ValOf $g.V), (SigOf $rows $g.R), (ValOf $g.R))
    }

    # ---- every case produced a row -----------------------------------------
    $missing = @()
    foreach ($g in $grid) {
        if (-not (RowOf $g.V)) { $missing += "$($g.T) ByVal" }
        if (-not (RowOf $g.R)) { $missing += "$($g.T) ByRef" }
    }
    Check 'every-type-ran-in-both-modes' ($missing.Count -eq 0) `
          ("missing: " + $(if ($missing.Count) { $missing -join ', ' } else { 'none' }) + " (of $($grid.Count * 2) cases)")

    # ---- THE INVARIANT: the two modes describe the same planted value ------
    $disagree = @()
    foreach ($g in $grid) {
        $a = ValOf $g.V; $b = ValOf $g.R
        if ($a -ne $b) { $disagree += ("{0}: ByVal='{1}' ByRef='{2}'" -f $g.T, $a, $b) }
    }
    Check 'byval-and-byref-report-the-same-value' ($disagree.Count -eq 0) `
          ("disagreements: " + $(if ($disagree.Count) { $disagree -join ' | ' } else { 'none' }))

    # ---- the planted value, where it is a stable literal --------------------
    # Object and Variant are checked separately; Date is a serial number whose
    # exact rendering is not the point of this file.
    $wrong = @()
    foreach ($g in $grid) {
        if (-not $g.Want) { continue }
        foreach ($m in @(@{ N='ByVal'; F=$g.V }, @{ N='ByRef'; F=$g.R })) {
            $got = ValOf $m.F
            if ($got -ne $g.Want) { $wrong += ("{0} {1}: got '{2}' want '{3}'" -f $g.T, $m.N, $got, $g.Want) }
        }
    }
    Check 'each-type-decodes-to-its-planted-value' ($wrong.Count -eq 0) `
          ("mismatches: " + $(if ($wrong.Count) { $wrong -join ' | ' } else { 'none' }))

    # ---- an object is an object in both modes ------------------------------
    Check 'an-object-parameter-decodes-in-both-modes' `
          (((ValOf 'VObject') -match '^[A-Za-z_][A-Za-z0-9_]*@0x[0-9A-F]+$') -and
           ((ValOf 'RObject') -match '^[A-Za-z_][A-Za-z0-9_]*@0x[0-9A-F]+$')) `
          ("ByVal='" + (ValOf 'VObject') + "' ByRef='" + (ValOf 'RObject') + "'")

    # ---- a Variant reads its held value, not a pointer ---------------------
    Check 'a-variant-parameter-reads-its-held-value-in-both-modes' `
          (((ValOf 'VVariant') -match '42') -and ((ValOf 'RVariant') -match '42')) `
          ("ByVal='" + (ValOf 'VVariant') + "' ByRef='" + (ValOf 'RVariant') + "'")

    # ---- THE TYPE CAME FROM THE P-CODE, not from the value validating -------
    #
    # A BSTR and a SAFEARRAY prove themselves whatever the opcode table knows,
    # so checking values alone would pass with the type table empty. The
    # signature is what says the walk recovered a type -- and `&` is what says
    # it recovered the ByRef one, which is a different opcode family.
    $noSig = @(); $noRef = @()
    foreach ($g in $grid) {
        if (-not $g.Sig) { continue }     # LongLong& is an untyped 8-byte ref
        if ((SigOf $rows $g.V) -notmatch [regex]::Escape($g.Sig)) { $noSig += ("{0} ByVal sig='{1}'" -f $g.T, (SigOf $rows $g.V)) }
        if ((SigOf $rows $g.R) -notmatch ([regex]::Escape($g.Sig) + '&')) { $noRef += ("{0} ByRef sig='{1}'" -f $g.T, (SigOf $rows $g.R)) }
    }
    Check 'byval-signatures-name-the-declared-type' ($noSig.Count -eq 0) `
          ("missing: " + $(if ($noSig.Count) { $noSig -join ' | ' } else { 'none' }))
    Check 'byref-signatures-name-the-declared-type-with-an-ampersand' ($noRef.Count -eq 0) `
          ("missing: " + $(if ($noRef.Count) { $noRef -join ' | ' } else { 'none' }))

    # ---- no case fell back to a bare raw qword ------------------------------
    # An undecoded slot is honest but is not a pass here: every one of these has
    # a declared type and a planted value, so a raw pointer means a gap.
    $raw = @()
    foreach ($g in $grid) {
        foreach ($m in @(@{ N='ByVal'; F=$g.V }, @{ N='ByRef'; F=$g.R })) {
            if ((ValOf $m.F) -match '^0x[0-9A-F]+$') { $raw += ("{0} {1}='{2}'" -f $g.T, $m.N, (ValOf $m.F)) }
        }
    }
    Check 'no-parameter-reads-as-an-undecoded-qword' ($raw.Count -eq 0) `
          ("raw: " + $(if ($raw.Count) { $raw -join ' | ' } else { 'none' }))

    # ---- THE FORMS THAT HAVE ONLY ONE MODE --------------------------------
    #
    # Not every parameter has both. `ByVal` on an ARRAY is a compile error
    # (MS-VBAL 5.3.1.5), a UDT parameter is always by reference, and a
    # ParamArray is always a resizable ByRef array of Variant. Asserting a
    # ByVal case for those would be asserting against the language.
    Write-Output ''
    Write-Output 'ByRef-only forms:'
    foreach ($f in @('RUdt','RArray','RParamArray')) {
        Write-Output ('  {0,-14} sig={1,-14} {2}' -f $f, (SigOf $rows $f), (ValOf $f))
    }
    Write-Output 'Optional:'
    foreach ($f in @('VOptional')) {
        $r = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq $f })
        for ($i = 0; $i -lt $r.Count; $i++) { Write-Output ('  {0,-14} {1}' -f "$f[$i]", $r[$i].args) }
    }

    # A record has no scalar value, so its address is the answer, the same shape an object gets.
    # It must be a plausible address: a stack address is 8-aligned and well above the first
    # page, while the record's packed members (X=55, Y=66 -> 0x4200000037) are neither.
    $udtVal = ValOf 'RUdt'
    $udtOk = $false
    if ($udtVal -match '^udt@0x([0-9A-F]+)$') {
        $addr = [Convert]::ToUInt64($Matches[1], 16)
        $udtOk = ($addr -gt 0x10000) -and (($addr % 8) -eq 0)
    }
    Check 'a-udt-parameter-reports-a-plausible-address' $udtOk `
          ("RUdt sig=" + (SigOf $rows 'RUdt') + " value='$udtVal' (must be 8-aligned, above the first page)")

    Check 'an-array-parameter-decodes-with-its-bounds' `
          ((ValOf 'RArray') -match 'Long\[1\.\.3\]\{111,222,333\}') `
          ("RArray sig=" + (SigOf $rows 'RArray') + " value='" + (ValOf 'RArray') + "'")

    Check 'a-paramarray-decodes-as-an-array-of-variants' `
          ((ValOf 'RParamArray') -match '\[0\.\.2\]|\[3\]') `
          ("RParamArray sig=" + (SigOf $rows 'RParamArray') + " value='" + (ValOf 'RParamArray') + "'")

    # An omitted Optional is the published `LitVar_Missing` marker, and the
    # supplied case must NOT read as one -- otherwise `Missing` would be what
    # the decoder says when it is unsure rather than when it is certain.
    $optRows = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'VOptional' } |
                 Sort-Object { [int]$_.span })
    if ($optRows.Count -ge 2) {
        Check 'a-supplied-optional-reads-its-value' `
              ([string]$optRows[0].args -match '=5$') `
              ("supplied: '" + $optRows[0].args + "'")
        Check 'an-omitted-optional-reads-its-default-not-missing' `
              ([string]$optRows[1].args -match '=99$') `
              ("omitted: '" + $optRows[1].args + "' (declared default 99)")
    } else {
        Check 'optional-ran-twice' $false "VOptional rows=$($optRows.Count), expected 2"
    }

    # ---- A TYPED VARIABLE INTO A VARIANT PARAMETER, and Decimal --------------
    #
    # ByRef into a Variant parameter is ordinary VBA and arrives as a VARIANT
    # tagged VT_BYREF; it read as a raw qword until the tag was followed.
    # Decimal is the one VBA value that lives only inside a Variant.
    Check 'a-typed-variable-byval-into-a-variant-reads-its-value' `
          ((ValOf 'VVarFromStr') -eq '"hello"') ("VVarFromStr='" + (ValOf 'VVarFromStr') + "'")
    Check 'a-string-byref-into-a-variant-reads-through-the-byref-tag' `
          ((ValOf 'RVarFromStr') -eq '"hello"') ("RVarFromStr='" + (ValOf 'RVarFromStr') + "'")
    Check 'a-long-byref-into-a-variant-reads-through-the-byref-tag' `
          ((ValOf 'RVarFromLng') -eq 'Long(123456)') ("RVarFromLng='" + (ValOf 'RVarFromLng') + "'")
    Check 'a-decimal-in-a-variant-reads-exactly' `
          ((ValOf 'VDec') -eq 'Decimal(1.5)') ("VDec='" + (ValOf 'VDec') + "'")
    Check 'a-decimal-in-a-byref-variant-reads-exactly' `
          ((ValOf 'RDec') -eq 'Decimal(1.5)') ("RDec='" + (ValOf 'RDec') + "'")


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} declared types, ByVal and ByRef, all agreeing" -f $grid.Count)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
