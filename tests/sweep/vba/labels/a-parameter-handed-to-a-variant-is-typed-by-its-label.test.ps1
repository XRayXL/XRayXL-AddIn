# A parameter the body only hands to a Variant -- a VBA Variant parameter, or a built-in such as
# Year() -- has no typed load of its own. VBA wraps it in a by-reference Variant whose VARTYPE the
# callee reads it by, and that label types it: the declared type, and the planted value. An array
# parameter only ReDimmed is labelled the same way, by the element type the ReDim allocates.
# Diagnostics are on (suite.psd1), so each signature names the opcode that typed it: 950 or 951 is
# the label, anything else a typed instruction in the body.
. (Join-Path $PSScriptRoot '..\..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\..\_xray_common.ps1')

$classCode = @'
Public Tag As Long
'@

$enumCode = @'
Public Enum XLColour
    XLRed = 1
    XLBlue = 2
End Enum
'@

$moduleCode = @'
Public gSink As Variant, gN As Long

' The one place each parameter goes: a Variant, by reference.
Public Sub Sink(v As Variant)
    If IsObject(v) Then Set gSink = v Else gSink = v
End Sub

' Each body hands p to Sink and does nothing else with it.
Public Sub VByte(ByVal p As Byte): Sink p: End Sub
Public Sub RByte(ByRef p As Byte): Sink p: End Sub
Public Sub VInteger(ByVal p As Integer): Sink p: End Sub
Public Sub RInteger(ByRef p As Integer): Sink p: End Sub
Public Sub VLong(ByVal p As Long): Sink p: End Sub
Public Sub RLong(ByRef p As Long): Sink p: End Sub
Public Sub VLongLong(ByVal p As LongLong): Sink p: End Sub
Public Sub RLongLong(ByRef p As LongLong): Sink p: End Sub
Public Sub VSingle(ByVal p As Single): Sink p: End Sub
Public Sub RSingle(ByRef p As Single): Sink p: End Sub
Public Sub VDouble(ByVal p As Double): Sink p: End Sub
Public Sub RDouble(ByRef p As Double): Sink p: End Sub
Public Sub VCurrency(ByVal p As Currency): Sink p: End Sub
Public Sub RCurrency(ByRef p As Currency): Sink p: End Sub
Public Sub VString(ByVal p As String): Sink p: End Sub
Public Sub RString(ByRef p As String): Sink p: End Sub
Public Sub VObject(ByVal p As Object): Sink p: End Sub
Public Sub RObject(ByRef p As Object): Sink p: End Sub
Public Sub VDate(ByVal p As Date): Sink p: End Sub
Public Sub RDate(ByRef p As Date): Sink p: End Sub
Public Sub VBoolean(ByVal p As Boolean): Sink p: End Sub
Public Sub RBoolean(ByRef p As Boolean): Sink p: End Sub
Public Sub VLongPtr(ByVal p As LongPtr): Sink p: End Sub
Public Sub RLongPtr(ByRef p As LongPtr): Sink p: End Sub
Public Sub VEnum(ByVal p As XLColour): Sink p: End Sub
Public Sub REnum(ByRef p As XLColour): Sink p: End Sub
Public Sub VClass(ByVal p As XLTag): Sink p: End Sub
Public Sub RClass(ByRef p As XLTag): Sink p: End Sub
Public Sub VCom(ByVal p As Collection): Sink p: End Sub
Public Sub RCom(ByRef p As Collection): Sink p: End Sub
' An array is ByRef only (MS-VBAL 5.3.1.5).
Public Sub RArray(ByRef p() As Long): Sink p: End Sub

' A built-in that takes a Variant, as Demo_AddMonths hands its date to Year().
Public Sub BVDate(ByVal p As Date): gN = Year(p): End Sub
Public Sub BRDate(ByRef p As Date): gN = Year(p): End Sub
Public Sub BVLong(ByVal p As Long): gN = IsNumeric(p): End Sub
Public Sub BRString(ByRef p As String): gN = IsNumeric(p): End Sub

' A ReDim of the array parameter and nothing else: the ReDim names the element type.
Public Sub RRedimD(ByRef p() As Double): ReDim p(1 To 3): End Sub
Public Sub RRedimL(ByRef p() As Long): ReDim Preserve p(1 To 3): End Sub
Public Sub RRedimS(ByRef p() As String): ReDim p(1 To 3): End Sub
Public Sub RRedimV(ByRef p() As Variant): ReDim Preserve p(1 To 3): End Sub
' The control: a Variant is ReDimmed by RedimVar, which names no element type.
Public Sub RRedimVar(ByRef p As Variant): ReDim p(1 To 3): End Sub

Public Sub Drive()
    Dim byt As Byte:      byt = 7
    Dim int16 As Integer: int16 = 1234
    Dim lng As Long:      lng = 123456
    Dim ll As LongLong:   ll = 4294967296#
    Dim sng As Single:    sng = 1.5
    Dim dbl As Double:    dbl = 2748.5
    Dim cur As Currency:  cur = 9.99
    Dim str1 As String:   str1 = "hello"
    Dim obj As Object:    Set obj = New XLTag
    Dim dt As Date:       dt = 43831
    Dim bl As Boolean:    bl = True
    Dim lp As LongPtr:    lp = 4294967296#
    Dim en As XLColour:   en = XLBlue
    Dim cls As XLTag:     Set cls = New XLTag: cls.Tag = 77
    Dim com As Collection: Set com = New Collection: com.Add 1
    Dim arr(1 To 3) As Long: arr(1) = 111: arr(2) = 222: arr(3) = 333

    VByte byt:        RByte byt
    VInteger int16:   RInteger int16
    VLong lng:        RLong lng
    VLongLong ll:     RLongLong ll
    VSingle sng:      RSingle sng
    VDouble dbl:      RDouble dbl
    VCurrency cur:    RCurrency cur
    VString str1:     RString str1
    VObject obj:      RObject obj
    VDate dt:         RDate dt
    VBoolean bl:      RBoolean bl
    VLongPtr lp:      RLongPtr lp
    VEnum en:         REnum en
    VClass cls:       RClass cls
    VCom com:         RCom com
    RArray arr
    BVDate dt:        BRDate dt
    BVLong lng:       BRString str1

    Dim rd() As Double:  ReDim rd(1 To 2): rd(1) = 1.5: rd(2) = 2.5
    Dim rl() As Long:    ReDim rl(1 To 2): rl(1) = 11: rl(2) = 22
    Dim rs() As String:  ReDim rs(1 To 2): rs(1) = "a": rs(2) = "b"
    Dim rv() As Variant: ReDim rv(1 To 2): rv(1) = 1.5: rv(2) = "x"
    Dim vv As Variant:   vv = 5
    RRedimD rd: RRedimL rl: RRedimS rs: RRedimV rv: RRedimVar vv
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'Labelled' @(
        @{ Kind=1; Name='LabelCase'; Code=$moduleCode }
        @{ Kind=2; Name='XLTag'; Code=$classCode }
        @{ Kind=1; Name='XLTypes'; Code=$enumCode }
    )
    $leaf = (Get-XRayMacroBook).Leaf

    [void](Set-XRayTraceParam $sx 'VBA' 'ARGS'  'TRUE')
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }
    $disarm = [string](Wait-LogLine $paths.Log 'VBA tracing: disarmed' $mark)

    $rows = @(Read-TraceRows $sx.ProcId)
    # the label stripped, address included
    function ValOf([string]$fn) {
        $r = EntryRowOf $rows $fn
        if ($null -eq $r) { return '(no row)' }
        return ([string]$r.args) -replace '^a1(:[^=]*)?=', ''
    }

    # Spelt as the loads spell them: Date is Double, Boolean Integer, a ByRef LongLong Ref&.
    # Load: VBA passes a ByVal String or object, and a ByRef class, on with a typed load (663, 664,
    # 744), which names it before any label.
    $object = '^[A-Za-z_][A-Za-z0-9_]*@0x[0-9A-F]+'
    $cases = @(
        @{ F='VByte';     Sig='Byte';      Want='7' }
        @{ F='RByte';     Sig='Byte&';     Want='7' }
        @{ F='VInteger';  Sig='Integer';   Want='1234' }
        @{ F='RInteger';  Sig='Integer&';  Want='1234' }
        @{ F='VLong';     Sig='Long';      Want='123456' }
        @{ F='RLong';     Sig='Long&';     Want='123456' }
        @{ F='VLongLong'; Sig='LongLong';  Want='4294967296' }
        @{ F='RLongLong'; Sig='Ref&';      Want='4294967296' }
        @{ F='VSingle';   Sig='Single';    Want='1.5' }
        @{ F='RSingle';   Sig='Single&';   Want='1.5' }
        @{ F='VDouble';   Sig='Double';    Want='2748.5' }
        @{ F='RDouble';   Sig='Double&';   Want='2748.5' }
        @{ F='VCurrency'; Sig='Currency';  Want='9.9900' }
        @{ F='RCurrency'; Sig='Currency&'; Want='9.9900' }
        @{ F='VString';   Sig='String';    Want='"hello"'; Load=$true }
        @{ F='RString';   Sig='String&';   Want='"hello"' }
        @{ F='VObject';   Sig='Object';    Match=$object; Load=$true }
        @{ F='RObject';   Sig='Object&';   Match=$object }
        @{ F='VDate';     Sig='Double';    Want='43831' }
        @{ F='RDate';     Sig='Double&';   Want='43831' }
        @{ F='VBoolean';  Sig='Integer';   Want='-1' }
        @{ F='RBoolean';  Sig='Integer&';  Want='-1' }
        @{ F='VLongPtr';  Sig='LongLong';  Want='4294967296' }
        @{ F='RLongPtr';  Sig='Ref&';      Want='4294967296' }
        @{ F='VEnum';     Sig='Long';      Want='2' }
        @{ F='REnum';     Sig='Long&';     Want='2' }
        @{ F='VClass';    Sig='Object';    Match=$object; Load=$true }
        @{ F='RClass';    Sig='Object&';   Match=$object; Load=$true }
        @{ F='VCom';      Sig='Object';    Match=$object; Load=$true }
        @{ F='RCom';      Sig='Object&';   Match=$object; Load=$true }
        @{ F='RArray';    Sig='Ref&';      Want='Long[1..3]{111,222,333}' }
        @{ F='BVDate';    Sig='Double';    Want='43831' }
        @{ F='BRDate';    Sig='Double&';   Want='43831' }
        @{ F='BVLong';    Sig='Long';      Want='123456' }
        @{ F='BRString';  Sig='String&';   Want='"hello"' }
        @{ F='RRedimD';   Sig='Ref&';      Want='Double[1..2]{1.5,2.5}' }
        @{ F='RRedimL';   Sig='Ref&';      Want='Long[1..2]{11,22}' }
        @{ F='RRedimS';   Sig='Ref&';      Want='String[1..2]{"a","b"}' }
        @{ F='RRedimV';   Sig='Ref&';      Want='Variant[1..2]{1.5,"x"}' }
    )

    # `Double#951|exit635`: the type, the opcode that named it, and the procedure's exit
    function TypeOf([string]$fn) { return (SigOf $rows $fn) -replace '#.*$', '' }
    function OpOf([string]$fn) { if ((SigOf $rows $fn) -match '#(\d+)') { return [int]$Matches[1] } else { return 0 } }

    Write-Output ''
    Write-Output ('  {0,-10} {1,-14} {2}' -f 'procedure', 'typetext', 'args')
    foreach ($c in $cases) { Write-Output ('  {0,-10} {1,-14} {2}' -f $c.F, (SigOf $rows $c.F), (ArgsOf $rows $c.F)) }

    $missing = @($cases | Where-Object { $null -eq (EntryRowOf $rows $_.F) } | ForEach-Object { $_.F })
    Check 'every-case-ran' ($missing.Count -eq 0) ("missing: " + $(if ($missing) { $missing -join ', ' } else { 'none' }))

    $badSig = @($cases | Where-Object { (TypeOf $_.F) -cne $_.Sig } |
                ForEach-Object { "{0}: '{1}' want '{2}'" -f $_.F, (SigOf $rows $_.F), $_.Sig })
    Check 'each-parameter-is-typed-as-declared' ($badSig.Count -eq 0) `
          ("wrong: " + $(if ($badSig) { $badSig -join ' | ' } else { 'none' }))

    $badVal = @()
    foreach ($c in $cases) {
        $v = ValOf $c.F
        $ok = if ($c.Match) { $v -match $c.Match } else { $v -ceq $c.Want }
        if (-not $ok) { $badVal += ("{0}: '{1}' want '{2}{3}'" -f $c.F, $v, $c.Want, $c.Match) }
    }
    Check 'each-value-reads-as-planted' ($badVal.Count -eq 0) `
          ("wrong: " + $(if ($badVal) { $badVal -join ' | ' } else { 'none' }))

    # Which cases the label typed, by the opcode each signature names; the disarm line must agree.
    # CDargRef, CVarRef, Redim, RedimPreserve
    $labelOps = @(950, 951, 1473, 1474)
    $byLabel = @($cases | Where-Object { (OpOf $_.F) -in $labelOps } | ForEach-Object { $_.F })
    $byLoad  = @($cases | Where-Object { (OpOf $_.F) -notin $labelOps } | ForEach-Object { "{0}#{1}" -f $_.F, (OpOf $_.F) })
    Write-Output ''
    Write-Output ("typed by the label ({0}): {1}" -f $byLabel.Count, ($byLabel -join ' '))
    Write-Output ("typed by a load ({0}):  {1}" -f $byLoad.Count, ($byLoad -join ' '))
    $wrongRoute = @($cases | Where-Object { ((OpOf $_.F) -in $labelOps) -eq [bool]$_.Load } |
                    ForEach-Object { "{0}: {1}" -f $_.F, (SigOf $rows $_.F) })
    Check 'the-label-types-exactly-what-has-no-load' ($wrongRoute.Count -eq 0) `
          ("wrong route: " + $(if ($wrongRoute) { $wrongRoute -join ' | ' } else { 'none' }))
    $logged = if ($disarm -match '(\d+) parameter\(s\) typed by a Variant label') { [int]$Matches[1] } else { 0 }
    Check 'the-disarm-line-counts-each-entry-once' ($logged -eq $byLabel.Count) `
          "disarm line $logged, signatures $($byLabel.Count)"
    # Demo_AddMonths's own shape: a ByVal Date handed to Year() has nothing but the label.
    Check 'a-date-handed-to-a-built-in-is-typed-by-its-label' ((OpOf 'BVDate') -eq 951) "BVDate [$(SigOf $rows 'BVDate')]"
    Check 'no-label-went-unread' ($disarm -notmatch '\(([1-9]\d*) label\(s\) fit no type\)') `
          ($(if ($disarm -match '\((\d+) label\(s\) fit no type\)') { "$($Matches[1]) label(s) fit no type" } else { 'none' }))

    # A Variant ReDims through RedimVar, which says nothing about an element type or an array.
    Check 'a-redimmed-variant-is-not-read-as-an-array' ((TypeOf 'RRedimVar') -ne 'Ref&') `
          "RRedimVar [$(SigOf $rows 'RRedimVar')] $(ArgsOf $rows 'RRedimVar')"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("{0} of {1} parameters typed by their label; BVDate [{2}]" -f $byLabel.Count, $cases.Count, (ArgsOf $rows 'BVDate'))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
