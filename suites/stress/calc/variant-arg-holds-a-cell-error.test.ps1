$case = @{ Name='variant-arg-holds-a-cell-error'
     Modules=@{
       'M'=@'
' AN EXCEL ERROR HANDED TO A VARIANT PARAMETER, which is what `=MyUdf(NA())`
' does, and what every sheet does the moment a #N/A reaches a UDF.
'
' TWO SHAPES, AND THEY ARE NOT THE SAME, which is the point of this case:
'
'   =TakeV("x", NA())   a VALUE     -- the parameter holds an error Variant
'   =RefV("x", A1)      a REFERENCE -- the parameter holds a RANGE OBJECT, and
'                                      VBA coerces through its default property,
'                                      so IsError() is True while the argument
'                                      itself is an object
'
' The trace reports what the parameter HOLDS. Reading VBA's own IsError/CStr as
' the oracle for the second shape is a mistake -- they answer about the coerced
' value, not about the argument -- and this case pins both answers so neither
' can be "corrected" into the other.
Public Function TakeV(ByVal tag As String, ByVal v As Variant) As Variant
    Saw tag, TypeName(v)
    TakeV = v
End Function

Public Function RefV(ByVal tag As String, ByVal v As Variant) As Variant
    Saw tag, TypeName(v)
    RefV = 1
End Function

Public Sub Saw(ByVal tag As String, ByVal vbaSays As String)
    Dim z As Long
    z = Len(vbaSays)
End Sub

' The error as a RETURN value, and then travelling on as another UDF's input.
Public Function GiveErr() As Variant
    GiveErr = CVErr(2042)
End Function

Public Function PassOn(ByVal v As Variant) As Variant
    PassOn = v
End Function
'@
     }
     Cells=@{
        'A1' = '=NA()'
        'B1' = '=TakeV("na",NA())'
        'B2' = '=TakeV("div0",1/0)'
        'B3' = '=TakeV("value",VALUE("x"))'
        'B4' = '=TakeV("num",SQRT(-1))'
        'C1' = '=RefV("ref",A1)'
        'D1' = '=GiveErr()'
        'D2' = '=PassOn(GiveErr())'
     }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $entries = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })
        $exits   = @($t.rows | Where-Object { $_.kind -eq 'exit'  -and $_.source -eq 'VBA' })

        $tagOf = { param($r) ($r.args -match '(?:^|\s)a1:[^=\s]*="([^"]*)"') | Out-Null; $Matches[1] }

        # ---- an error VALUE in a Variant parameter -------------------------
        $want = @{ 'na' = '#N/A'; 'div0' = '#DIV/0!'; 'value' = '#VALUE!'; 'num' = '#NUM!' }
        $seen = @{}
        foreach ($r in @($entries | Where-Object { $_.function -eq 'TakeV' })) {
            $seen[(& $tagOf $r)] = ArgOf $r 2
        }
        foreach ($k in $want.Keys) {
            if (-not $seen.ContainsKey($k)) { return "TakeV('$k') never traced" }
            # THE FAILURE THIS EXISTS FOR. `Missing` and an Excel error are BOTH
            # VT_ERROR, told apart only by the exact SCODE -- an omitted Optional
            # is DISP_E_PARAMNOTFOUND. Loosen that and every #N/A in a sheet
            # reads as "parameter omitted": a supplied argument reported as an
            # absent one.
            if ($seen[$k] -eq 'Missing') {
                return "$k : an Excel error rendered as 'Missing' -- a supplied argument reported as omitted" }
            if ($seen[$k] -ne $want[$k]) {
                return "$k : argument rendered [$($seen[$k])], expected $($want[$k])" }
        }

        # ---- a REFERENCE is an object, and must not be "fixed" into a value --
        $ref = @($entries | Where-Object { $_.function -eq 'RefV' })
        if ($ref.Count -eq 0) { return 'RefV never traced' }
        $refArg = ArgOf $ref[0] 2
        # THE DISTINCTION IS SHARPER NOW, not gone. With OBJECTS on, the Range is
        # described rather than shown as an address -- so the two shapes read:
        #
        #   =TakeV("na", NA())   ->  #N/A                      the parameter HOLDS the error
        #   =RefV("ref", A1)     ->  Range(...!A1)=#N/A        it holds a RANGE whose value is
        #
        # Both are `IsError` to VBA. Only one of them is an error value, and the
        # row must not conflate them.
        if ($refArg -notmatch "^Range@0x[0-9A-F]+\('?\[") {
            return "a cell REFERENCE passed to a Variant parameter rendered [$refArg]; Excel hands a Range object there, and the trace must say what the parameter holds" }
        if ($refArg -notmatch '#N/A') {
            return "the referenced cell holds #N/A, so the Range's value should say so: [$refArg]" }
        # VBA agrees, in its own words, through TypeName rather than a coercion.
        $saw = @{}
        foreach ($r in @($entries | Where-Object { $_.function -eq 'Saw' })) { $saw[(& $tagOf $r)] = ArgOf $r 2 }
        if ($saw['ref'] -ne '"Range"') {
            return "VBA called the reference argument $($saw['ref']), not Range -- the fixture no longer tests what it claims" }
        if ($saw['na'] -ne '"Error"') {
            return "VBA called the value argument $($saw['na']), not Error -- the fixture no longer tests what it claims" }

        # ---- the same error as a RETURN value ------------------------------
        $give = @($exits | Where-Object { $_.function -eq 'GiveErr' })
        if ($give.Count -eq 0) { return 'GiveErr never traced' }
        if ([string]$give[0].ret -ne '#N/A') {
            return "a Variant return of CVErr(2042) rendered [$([string]$give[0].ret)], expected #N/A" }

        # ---- and travelling on, as a second UDF's argument -----------------
        $pass = @($entries | Where-Object { $_.function -eq 'PassOn' })
        if ($pass.Count -eq 0) { return 'PassOn never traced -- the error did not travel on' }
        if ((ArgOf $pass[0] 1) -ne '#N/A') {
            return "PassOn received [$(ArgOf $pass[0] 1)], expected #N/A" }
        $null }
     Why='an Excel error reaching a Variant -- as an argument, as a return, and
          passed on -- spelt as Excel spells it; and the reference case beside it,
          where the parameter holds a Range and VBA coerces' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
