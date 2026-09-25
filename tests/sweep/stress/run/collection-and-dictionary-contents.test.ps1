$case = @{ Name='collection-and-dictionary-contents'
     Modules=@{
       'M'=@'
' Each shape of Collection and Dictionary crosses one traced boundary, tagged, so the trace
' can be checked against what this macro put in them.
Public Sub Take(ByVal tag As String, ByVal v As Variant)
End Sub

Public Sub TakeColl(ByVal tag As String, ByVal c As Collection)
End Sub

Public Sub Go()
    Dim c As Collection, inner As Collection, d As Object, none As Object, i As Long

    Set c = New Collection
    c.Add 1.5: c.Add "x": c.Add Nothing: c.Add Empty: c.Add Array(1, 2): c.Add CLng(7)
    Take "mixed", c

    Take "emptycoll", New Collection

    Set inner = New Collection
    inner.Add "i"
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a", 1
    d.Add 2, "two"
    d.Add "n", Nothing
    d.Add "c", inner
    Take "dict", d

    Take "emptydict", CreateObject("Scripting.Dictionary")

    Set c = New Collection
    c.Add c
    Take "self", c
    c.Remove 1

    Set c = New Collection
    c.Add ThisWorkbook.Worksheets("S1").Range("B2")
    Take "range", c

    Set c = New Collection
    For i = 1 To 5000: c.Add i: Next
    Take "big", c

    TakeColl "nothingcoll", Nothing
    Take "nothingvar", none
End Sub
'@
     }
     Cells=@{ 'B2' = '7' }
     Trigger=@{ Kind='Run'; Name='Go' }
     Expect={ param($t)
        $got = @{}
        foreach ($r in @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' -and
                                                  ($_.function -eq 'Take' -or $_.function -eq 'TakeColl') })) {
            $got[(ArgOf $r 1).Trim('"')] = ArgOf $r 2
        }
        $plain = { param($s) $s -replace '@0x[0-9A-F]+', '@0x*' }

        $want = [ordered]@{
            'mixed'       = 'Collection@0x*=Variant[1..6]{1.5,"x",Nothing,Empty,Variant[0..1]{Integer(1),Integer(2)},Long(7)}'
            'emptycoll'   = 'Collection@0x*=Variant[1..0]{}'
            'dict'        = 'Dictionary@0x*=Variant[0..3,0..1]{{"a",Integer(1)},{Integer(2),"two"},{"n",Nothing},{"c",Collection@0x*=Variant[1..1]{"i"}}}'
            'emptydict'   = 'Dictionary@0x*=Variant[0..-1,0..1]{}'
            'big'         = 'Collection@0x*'
            'nothingcoll' = 'Nothing'
            'nothingvar'  = 'Nothing'
        }
        foreach ($k in $want.Keys) {
            if (-not $got.ContainsKey($k)) { return "$k : never traced" }
            $g = & $plain $got[$k]
            if ($g -cne $want[$k]) { return "$k : got [$($got[$k])], want [$($want[$k])]" }
        }

        # A Collection holding itself is walked once: the inner one is named, at the same address.
        if ($got['self'] -cnotmatch '^Collection@0x([0-9A-F]+)=Variant\[1\.\.1\]\{Collection@0x([0-9A-F]+)\}$' -or
            $Matches[1] -ne $Matches[2]) {
            return "self : got [$($got['self'])]" }

        # An object inside a container is described as it would be on its own.
        if ($got['range'] -cnotmatch '^Collection@0x[0-9A-F]+=Variant\[1\.\.1\]\{Range@0x[0-9A-F]+\(.*!B2\)=7\}$') {
            return "range : got [$($got['range'])]" }
        $null }
     Why='a Collection and a Dictionary show their contents in the array grammar: every item,
          Nothing and Empty among them, a nested container and a Range inside one, a container
          that holds itself walked once, and one over the item ceiling named but not read' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
