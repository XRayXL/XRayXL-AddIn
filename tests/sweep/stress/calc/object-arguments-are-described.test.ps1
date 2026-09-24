$case = @{ Name='object-arguments-are-described'
     Modules=@{
       'M'=@'
' A Range, a Worksheet and a Workbook reaching a VBA UDF, which is what
' `=MyUdf(A1)` does and what every sheet does the moment a UDF takes a range.
'
' WHAT EXCEL HANDS BACK FOR Value2 IS MEASURED HERE, NOT ASSUMED. A single cell,
' a single row, a single column and a block are four different shapes and Excel
' coerces between them by rules that are easy to misremember -- notably that a
' one-row range is a 2-D array of 1 x N and not a 1-D array. Whatever it does,
' the array renderer prints BOTH bounds of every dimension, so this test records
' the truth rather than a guess about it.
'
' VBA'S OWN TypeName() RIDES IN THE TRACE beside ours, as an oracle: `Saw` is a
' traced procedure taking a String, so the trace holds what the tracer called the
' object and what VBA called it, side by side.
Public Function TakeObj(ByVal tag As String, ByVal v As Variant) As Variant
    Saw tag, TypeName(v)
    TakeObj = 1
End Function

Public Sub Saw(ByVal tag As String, ByVal vbaSays As String)
    Dim z As Long
    z = Len(vbaSays)
End Sub

' A Worksheet and a Workbook, which cannot come from a cell reference.
Public Function TakeSheet(ByVal tag As String) As Variant
    Saw tag, TypeName(Application.ActiveSheet)
    TakeObj2 tag, Application.ActiveSheet
    TakeSheet = 1
End Function

Public Function TakeBook(ByVal tag As String) As Variant
    Saw tag, TypeName(Application.ActiveWorkbook)
    TakeObj2 tag, Application.ActiveWorkbook
    TakeBook = 1
End Function

Public Sub TakeObj2(ByVal tag As String, ByVal v As Variant)
    Dim z As Long
    z = 1
End Sub

' A class Excel does not own, to prove an unknown object is still NAMED.
Public Function TakeColl(ByVal tag As String) As Variant
    Dim c As New Collection
    c.Add 1
    Saw tag, TypeName(c)
    TakeObj2 tag, c
    TakeColl = 1
End Function
'@
     }
     Cells=@{
        'A1' = '11'
        'B1' = '12'
        'C1' = '13'
        'A2' = '21'
        'B2' = '22'
        'C2' = '23'
        'E1' = '=TakeObj("cell",A1)'
        'E2' = '=TakeObj("row",A1:C1)'
        'E3' = '=TakeObj("col",A1:A2)'
        'E4' = '=TakeObj("block",A1:C2)'
        'E5' = '=TakeObj("wholecol",A:A)'
        'E6' = '=TakeSheet("sheet")'
        'E7' = '=TakeBook("book")'
        'E8' = '=TakeColl("coll")'
     }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        $entries = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })
        $tagOf = { param($r) if ($r.args -match '(?:^|\s)a1:[^=\s]*="([^"]*)"') { $Matches[1] } else { '' } }

        # What the tracer said and what VBA said, per tag.
        $ours = @{}
        foreach ($r in @($entries | Where-Object { $_.function -eq 'TakeObj' -or $_.function -eq 'TakeObj2' })) {
            $ours[(& $tagOf $r)] = ArgOf $r 2
        }
        $vba = @{}
        foreach ($r in @($entries | Where-Object { $_.function -eq 'Saw' })) {
            $vba[(& $tagOf $r)] = (ArgOf $r 2).Trim('"')
        }

        foreach ($tag in @('cell','row','col','block','wholecol','sheet','book','coll')) {
            if (-not $ours.ContainsKey($tag)) { return "$tag : never traced" }
            if ($ours[$tag] -eq '')           { return "$tag : no argument rendered" }
            # Nothing may still be a bare address: that is the answer when the setting is off or
            # a call failed, and with OBJECTS on and a live Excel neither is true.
            if ($ours[$tag] -match '^object@0x') {
                return "$tag : rendered a bare address [$($ours[$tag])] -- OBJECTS produced nothing" }
        }

        # ---- the three we describe ----------------------------------------
        foreach ($tag in @('cell','row','col','block','wholecol')) {
            # Excel quotes the prefix when the name needs it: this is Excel's own Address output,
            # not the tracer's `callerref`, so it is accepted as it comes.
            if ($ours[$tag] -notmatch "^Range@0x[0-9A-F]+\('?\[[^\]]+\][^!]+'?!") {
                return "$tag : expected Range([Book]Sheet!...), got [$($ours[$tag])]" }
        }
        if ($ours['sheet'] -notmatch '^Worksheet@0x[0-9A-F]+\(') { return "sheet : got [$($ours['sheet'])]" }
        if ($ours['book']  -notmatch '^Workbook@0x[0-9A-F]+\(')  { return "book : got [$($ours['book'])]" }

        # ---- an unknown class is still named -------------------------------
        # The class name is what VBA's TypeName says; the address follows it and
        # is ours, so compare only the part that is a claim about the class.
        $collCls = ($ours['coll'] -split '@')[0]
        if ($collCls -ne $vba['coll']) {
            return "coll : tracer says [$collCls], VBA's TypeName says [$($vba['coll'])]" }
        if ($ours['coll'] -notmatch '@0x[0-9A-F]+$') {
            return "coll : lost its address [$($ours['coll'])] -- an object is followed by that, not by its class" }

        # ---- the ceiling: a whole column is addressed, never read -----------
        if ($ours['wholecol'] -match '\)=') {
            return "wholecol : read the contents of a whole column [$($ours['wholecol'])]" }

        # ---- and the shapes Excel actually produced -------------------------
        # Recorded, not asserted against a guess: each must carry a value, and a multi-cell range both
        # bounds of every dimension; the detail line shows what they are.
        foreach ($tag in @('cell','row','col','block')) {
            if ($ours[$tag] -notmatch '\)=') { return "$tag : no value read, only an address [$($ours[$tag])]" }
        }
        foreach ($tag in @('row','col','block')) {
            if ($ours[$tag] -notmatch '=\w+\[-?\d+\.\.-?\d+,-?\d+\.\.-?\d+\]\{') {
                return "$tag : expected a 2-D array with both bounds, got [$($ours[$tag])]" }
        }
        # Expect returns its result, so a stray Write-Output here would become the return value
        # instead of the verdict.
        $null }
     Why='a Range, Worksheet, Workbook and an unowned class reaching a VBA UDF:
          each described, the unknown one still named, a whole column addressed
          but not read, and the shape Excel returns for Value2 measured rather
          than assumed' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
