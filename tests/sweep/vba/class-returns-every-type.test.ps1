# RETURN VALUES OUT OF A CLASS MODULE, ONE PER TYPE.
#
# Class and form procedures leave through exit opcodes the return-value work
# never mapped, so every one of them silently reported no result -- a
# Property Get and a class Function both read ret=''. The tracer now COUNTS that
# case (`returnsUnmapped`) and names the opcodes, which is how it was found:
# opcode 504 four times and 1664 three times, against five Subs and three
# Functions.
#
# WHICH IS NOT ENOUGH TO MAP THEM. All three Functions in that run returned
# Long, so the evidence cannot distinguish "1664 means Long" from "1664 is the
# generic class-Function exit and the type is somewhere else". Mapping it to
# Long on that basis would decode a String as an integer and print a confident
# wrong value -- the failure this project ranks worst, and the exact trap slot
# 634 already documents (it serves both LongLong and every typed array, and only
# the STORE opcode separates them).
#
# So this exercises one Function per type and reports the opcode each uses. If
# they differ, the opcodes are type-specific and can be mapped directly. If they
# are all the same, the type must come from somewhere else and the mapping is a
# different piece of work -- which is a result either way, and the point of
# running it before writing the table.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')

$classCode = @'
Public Function RByte() As Byte
    RByte = 7
End Function
Public Function RInteger() As Integer
    RInteger = 1234
End Function
Public Function RLong() As Long
    RLong = 123456
End Function
Public Function RSingle() As Single
    RSingle = 1.5
End Function
Public Function RDouble() As Double
    RDouble = 2748.5
End Function
Public Function RCurrency() As Currency
    RCurrency = 9.99
End Function
Public Function RString() As String
    gStringCalls = gStringCalls + 1
    RString = "hello"
End Function
Public Function RBoolean() As Boolean
    RBoolean = True
End Function
Public Function RVariant() As Variant
    RVariant = 42
End Function
' A Variant holding an ARRAY stores through the SAME opcode as a scalar one
' (measured in isolation), so nothing in the mapping can tell them apart --
' the VARIANT's own type tag does, which is DescribeVariantAt's job.
Public Function RVariantArr() As Variant
    RVariantArr = Array(11, 22, 33)
End Function
Public Function RVariantStr() As Variant
    RVariantStr = "in a variant"
End Function
' A TYPED ARRAY RESULT, which is the class path's own case: the class exit
' opcode 1664 carries no type, so only the STORE can name this, and the store is
' 671. Before that mapping existed this decoded as Unknown and was counted as
' returnsUnmapped.
Public Function RLongArr() As Long()
    Dim a(0 To 2) As Long
    a(0) = 11
    RLongArr = a
End Function
' THE SAME RETURN FROM A BUSIER BODY. Identical declaration and identical
' assignment to the result -- the only difference is three element stores
' instead of one. If this decodes and RLongArr does not, or the reverse, the
' fault is in the store SCAN and not in the mapping, because the mapping cannot
' see the difference.
Public Function RLongArrBusy() As Long()
    Dim b(0 To 2) As Long
    b(0) = 11
    b(1) = 22
    b(2) = 33
    RLongArrBusy = b
End Function
Public Function RObject() As Object
    Set RObject = Nothing
End Function
Public Sub RSub()
    Dim z As Long
    z = 1
End Sub
Public Property Get PGet() As Long
    PGet = 99
End Property
'@

$moduleCode = @'
' VBA COUNTS ITS OWN ACTIVATIONS, so the tracer can be checked against truth
' rather than against a reading of the tracer.
Public gStringCalls As Long
Public Function XRGetStringCalls() As Long
    XRGetStringCalls = gStringCalls
End Function

Public Sub R_Drive()
    Dim o As New XRRet
    Dim b As Byte, i As Integer, l As Long
    Dim s As Single, d As Double, c As Currency
    Dim t As String, v As Variant, x As Long
    Dim bo As Boolean, ob As Object
    Dim va As Variant, vs As Variant
    b = o.RByte
    i = o.RInteger
    l = o.RLong
    s = o.RSingle
    d = o.RDouble
    c = o.RCurrency
    t = o.RString
    bo = o.RBoolean
    v = o.RVariant
    va = o.RVariantArr
    vs = o.RVariantStr
    Dim la() As Long
    la = o.RLongArr
    la = o.RLongArrBusy
    Set ob = o.RObject
    o.RSub
    x = o.PGet
    Set o = Nothing
End Sub
'@

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $paths = Get-XRayPaths $sx.ProcId

    New-XRayMacroBook $sx 'ClsRet' @(
        @{ Kind=1; Name='RetCase'; Code=$moduleCode }
        @{ Kind=2; Name='XRRet'; Code=$classCode }
    )
    $book = Get-XRayMacroBook
    $leaf = $book.Leaf

    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $mark = Get-LogLength $paths.Log
    [void](Invoke-XRayCommand $sx 'XRayXL_Arm')
    $armLine = Wait-LogLine $paths.Log 'VBA tracing: ' $mark
    if ($armLine -notmatch 'ARMED') { Complete-Test -Fail -Detail "did not arm: $armLine" }

    $app.Run($leaf + '!R_Drive') | Out-Null
    $lossy = Stop-XRayTrace $sx
    if ($lossy) { Complete-Test -Fail -Detail $lossy }

    $totals = Wait-LogLine $paths.Log 'VBA trace: statements=' $mark 20
    $rows  = @(Read-TraceRows $sx.ProcId)
    $exits = @($rows | Where-Object { ($_.kind -eq 'exit' -and $_.source -eq 'VBA') })

    # Every typed Function must have run, or the report below is about nothing.
    $want = @('RByte','RInteger','RLong','RSingle','RDouble','RCurrency',
              'RString','RBoolean','RVariant','RVariantArr','RVariantStr',
              'RLongArr','RLongArrBusy','RObject','RSub','PGet')
    $seen = @($exits | ForEach-Object { $_.function })
    $missing = @($want | Where-Object { $seen -notcontains $_ })
    Check 'every-typed-function-ran' ($missing.Count -eq 0) `
          ("missing: " + ($missing -join ',') + " | saw: " + ($seen -join ','))

    # ---- WHAT THE TRACER CURRENTLY REPORTS -------------------------------
    Write-Output ''
    Write-Output 'returns by declared type:'
    # EVERY row per procedure, not just the first. Several of these produce TWO
    # exit rows and the first one showed an empty result -- which is a fact
    # about WHEN the result exists, and invisible if only one row is printed.
    foreach ($w in $want) {
        $r = @($exits | Where-Object { $_.function -eq $w })
        if (-not $r.Count) { Write-Output ("  {0,-12} (no exit row)" -f $w); continue }
        for ($i = 0; $i -lt $r.Count; $i++) {
            $note = if ($r[$i].trust) { [string]$r[$i].trust } else { '?' }
            Write-Output ("  {0,-12} [{1}] ret='{2}' rettype='{3}' closed={4}" -f `
                $w, $i, $r[$i].ret, $r[$i].rettype, $note)
        }
    }

    $unmapped = [regex]::Matches($totals, 'unmappedExitOp\[(\d+)\]=(\d+)')
    Write-Output ''
    Write-Output 'unmapped exit opcodes seen:'
    foreach ($m in $unmapped) { Write-Output ("  opcode {0} x{1}" -f $m.Groups[1].Value, $m.Groups[2].Value) }
    $distinct = @($unmapped | ForEach-Object { $_.Groups[1].Value })
    Write-Output ("=> {0} distinct unmapped opcode(s) across {1} declared types" -f $distinct.Count, $want.Count)

    # ---- THE ASSERTION: a value is decoded, or the gap is NAMED ----------
    # Not "every type decodes" -- that is the goal, and asserting it before the
    # mapping exists would just be a red. What must hold NOW is that a return
    # the tracer cannot decode is COUNTED with its opcode, so the gap can be
    # closed rather than merely noticed.
    # ---- A SUB HAS NO RESULT, AND MUST NOT BE GIVEN ONE ------------------
    #
    # The negative control, and it caught a real bug on its first run. The
    # type now comes from the instruction that STORED to [R14-8] -- but in a
    # Sub that slot is a LOCAL, not a return. RSub's body is `Dim z As Long:
    # z = 1`, and it was reported as ret='1' rettype='Long': a value that does
    # not exist, printed with a type. Exit opcode 504 is the class Sub exit and
    # now maps to None, so the store is never consulted for one.
    # ---- DOES VBA REALLY CALL IT TWICE, OR DO WE OPEN TWO FRAMES? --------
    #
    # RString, RVariant and RObject each produced TWO entry/exit pairs with
    # distinct spans, the first carrying an empty result. Properly paired, so
    # not a double-close -- but two frames for ONE activation would be a frame
    # -accounting bug, while two real calls would be VBA's own behaviour. Only
    # VBA can settle which, so it counts its own.
    $vbaCalls = [int]($app.Run($leaf + '!XRGetStringCalls'))
    $strEntries = @($rows | Where-Object { ($_.kind -eq 'entry' -and $_.source -eq 'VBA') -and $_.function -eq 'RString' }).Count
    Check 'activation-count-matches-vba' ($strEntries -eq $vbaCalls) `
          "VBA counted $vbaCalls call(s) to RString; the trace has $strEntries entry row(s)"

    $subRows = @($exits | Where-Object { $_.function -eq 'RSub' })
    $subWithRet = @($subRows | Where-Object { $_.ret })
    Check 'a-sub-reports-no-return-value' `
          (($subRows.Count -ge 1) -and ($subWithRet.Count -eq 0)) `
          ("RSub rows: " + (@($subRows | ForEach-Object { "ret='$($_.ret)' type='$($_.rettype)'" }) -join ' | '))

    # ---- THE TYPES THAT ARE MEASURED MUST DECODE, AND CORRECTLY ----------
    # The store/load pairs are measured in vbapcode_tables.h; Boolean reads as Integer
    # (-1/0) by the row model. Values as well as types, because a decoder that
    # reads the right slot with the wrong width still produces a number.
    $expect = @{
        RByte     = @('7',      'Byte')
        RInteger  = @('1234',   'Integer')
        RLong     = @('123456', 'Long')
        RSingle   = @('1.5',    'Single')
        RDouble   = @('2748.5', 'Double')
        # EXACT, four decimals: Currency is an integer scaled by 10,000 and is
        # now rendered in integer arithmetic by both columns. The old
        # `%.10g` could not hold its 19 significant digits.
        RCurrency = @('9.9900', 'Currency')
        RBoolean  = @('-1',     'Integer')
        PGet      = @('99',     'Long')
        RString   = @('"hello"', 'String')
    }
    $wrong = @()
    foreach ($k in $expect.Keys) {
        $r = @($exits | Where-Object { $_.function -eq $k })
        if ($r.Count -eq 0) { $wrong += "$k missing"; continue }
        if ($r[0].ret -ne $expect[$k][0] -or $r[0].rettype -ne $expect[$k][1]) {
            $wrong += ("{0} got '{1}'/'{2}' want '{3}'/'{4}'" -f $k, $r[0].ret, $r[0].rettype, $expect[$k][0], $expect[$k][1])
        }
    }
    Check 'measured-types-decode-with-the-right-value' ($wrong.Count -eq 0) `
          ($wrong -join ' ; ')

    # ---- VARIANT, INCLUDING WHEN IT HOLDS AN ARRAY ------------------------
    #
    # A Variant holding an Array stores through the SAME opcode as a scalar one,
    # so the store cannot distinguish them and the mapping must not try. The
    # VARIANT's own type tag is what knows: an array must read as an array, a
    # string in a Variant as a string.
    $vScalar = @($exits | Where-Object { $_.function -eq 'RVariant' })
    $vArray  = @($exits | Where-Object { $_.function -eq 'RVariantArr' })
    $vString = @($exits | Where-Object { $_.function -eq 'RVariantStr' })
    Check 'variant-scalar-decodes' `
          (($vScalar.Count -ge 1) -and ($vScalar[0].ret -eq '42')) `
          ("RVariant ret='$($vScalar[0].ret)' type='$($vScalar[0].rettype)'")
    Check 'variant-holding-an-array-reads-as-an-array' `
          (($vArray.Count -ge 1) -and ($vArray[0].ret -eq 'Variant[0..2]{11,22,33}')) `
          ("RVariantArr ret='$($vArray[0].ret)' type='$($vArray[0].rettype)'")
    Check 'variant-holding-a-string-reads-as-a-string' `
          (($vString.Count -ge 1) -and ($vString[0].ret -match 'in a variant')) `
          ("RVariantStr ret='$($vString[0].ret)' type='$($vString[0].rettype)'")

    # ---- OBJECT -----------------------------------------------------------
    # Nothing is a real answer for an object: `Set RObject = Nothing`.
    $obj = @($exits | Where-Object { $_.function -eq 'RObject' })
    Check 'object-decodes' `
          (($obj.Count -ge 1) -and ($obj[0].rettype -eq 'Object')) `
          ("RObject ret='$($obj[0].ret)' type='$($obj[0].rettype)'")

    $withRet = @($exits | Where-Object { $_.ret })
    $nUnmapped = 0
    if ($totals -match 'returnsUnmapped=(\d+)') { $nUnmapped = [int]$Matches[1] }
    # A Sub is neither decoded nor unmapped -- it HAS no result, and that is a
    # third correct outcome rather than a gap. Counting it as one would demand
    # a return value from the one procedure that cannot have one, so the case
    # could only ever be satisfied by the decoder inventing something.
    $subs = @($want | Where-Object { $_ -eq 'RSub' }).Count
    Check 'nothing-is-silently-undecoded' `
          (($withRet.Count + $nUnmapped + $subs) -ge $want.Count) `
          "decoded=$($withRet.Count) unmapped=$nUnmapped subs=$subs over $($want.Count) procedures"


    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("decoded {0} of {1}; {2} distinct unmapped opcode(s)" -f `
        $withRet.Count, $want.Count, $distinct.Count)
}
catch {
    Complete-Test -Fail -Detail ("exception: " + $_.Exception.Message)
}
