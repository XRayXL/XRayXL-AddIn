$case = @{ Name='range-of-mixed-types'
     Modules=@{
       'M'=@'
' A RANGE HOLDING A STRING, A NUMBER AND A DATE, handed to a UDF. A real sheet
' is not a block of Doubles, and each of those three reaches the decoder as a
' different VARTYPE inside one array.
'
' NO SCAFFOLDING. The UDF RETURNS its own oracle -- VBA's TypeName for the thing
' it was handed -- so that lands in the `ret` column of the very row under test.
' An earlier version called a helper procedure to get VBA's account into the
' trace, and the helper's own entry and exit rows then outnumbered the subject
' twenty to eight.
Public Function TakeMixed(ByVal v As Variant) As Variant
    TakeMixed = TypeName(v)
End Function
'@
     }
     Cells=@{
        'A1' = 'hello'
        'A2' = '42.5'
        'A3' = '=DATE(2026,9,12)'
        'C1' = '=TakeMixed(A1:A3)'
     }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $e = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'TakeMixed' })
        $x = @($t.rows | Where-Object { $_.kind -eq 'exit'  -and $_.function -eq 'TakeMixed' })
        if ($e.Count -eq 0 -or $x.Count -eq 0) { return 'TakeMixed was not traced on both sides' }

        # VBA's own account, from the same call's return value.
        if ([string]$x[0].ret -ne '"Range"') {
            return "VBA called the argument $([string]$x[0].ret), not Range -- the case no longer tests what it claims" }

        $a = [string]$e[0].args
        if ($a -notmatch '^a1:[^=]*=(.+)$') { return "no argument rendered: [$a]" }
        $got = $Matches[1]

        if ($got -notmatch "^Range@0x[0-9A-F]+\('?\[") {
            return "expected a described Range, got [$got]" }
        if ($got -notmatch '\)=(.+)$') { return "the range carried no value: [$got]" }
        $val = $Matches[1]

        # A column of three is 2-D, three rows by one column, and a mixed column is the shape a reader
        # is most likely to meet. One brace level per row.
        if ($val -notmatch '^Variant\[1\.\.3,1\.\.1\]\{\{(.+)\}\}$') {
            return "expected Variant[1..3,1..1]{{...},{...},{...}}, got [$val]" }
        $elems = $Matches[1] -split '\},\{'

        # Each type keeps its own rendering inside the one array; nothing coerces them.
        if ($elems.Count -ne 3)      { return "expected 3 elements, got [$($Matches[1])]" }
        if ($elems[0] -ne '"hello"') { return "the string element reads [$($elems[0])]" }
        if ($elems[1] -ne '42.5')    { return "the number element reads [$($elems[1])]" }

        # The date against Excel's own Value2, read after disarm: the property the tracer read. Value2
        # is the serial, so a date rendering would mean the decoder started interpreting.
        if (-not $t.cells.ContainsKey('A3')) { return 'the driver did not read A3' }
        $serial = [string]$t.cells['A3']
        if ($elems[2] -ne $serial) {
            return "Value2 gives a SERIAL, not a date: Excel says A3 is '$serial'; the trace says [$($elems[2])]" }
        if ($elems[2] -match '[/:-]') {
            return "the date element [$($elems[2])] looks formatted -- Value2 must report the underlying number" }
        $null }
     Why='a range of a string, a number and a date in one UDF argument: each type
          keeps its own rendering inside the array, and Value2 hands the date back
          as a SERIAL, checked against the same property read from the sheet' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
