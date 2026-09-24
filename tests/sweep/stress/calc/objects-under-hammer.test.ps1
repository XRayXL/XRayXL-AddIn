$case = @{ Name='objects-under-hammer'
     Modules=@{
       'M'=@'
' THE HAMMER FOR THE OBJECT DESCRIBER. Describing an object is the one thing
' this tracer does that CALLS EXCEL rather than reading it -- GetIDsOfNames,
' Invoke, QueryInterface, GetTypeInfo -- on the calculating thread, from inside
' VBE7's dispatch, at a statement boundary. Every one of those is a place a new
' failure can live, and none of them is exercised by describing three objects
' once.
'
' EVERY CALL CROSSES EVERY PATH, because they fail differently:
'
'   r      a Range of three cells   -- Address, Count, Value2 -> a SAFEARRAY
'                                      that has to be freed every time
'   big    a whole column           -- Count checked, Value2 NEVER asked for
'   sheet  a Worksheet              -- Name, then Parent, which is another object
'   book   a Workbook               -- Name
'   coll   a Collection             -- QueryInterface says no to all three, then
'                                      IProvideClassInfo names it
'   noObj  Nothing                  -- the path that must call nothing at all
'
' WHAT IS ON TRIAL is not the values -- the sibling cases assert those. It is
' whether calling Excel from a hook LEAKS, FAULTS or DEGRADES when it is done a
' great many times: a BSTR or VARIANT not freed, a reference not released, a
' guarded call that faults and trips the circuit breaker. None of that shows up
' in a single call.
Public Function HAll(ByVal n As Double, ByVal r As Variant, ByVal big As Variant) As Double
    ' NOT `empty`: that is a VBA RESERVED WORD, and declaring it is a syntax
    ' error that stops the module compiling -- which presents as the VBE
    ' opening in break mode and the case timing out, not as a failed assertion.
    Dim coll As New Collection
    Dim noObj As Object
    coll.Add 1

    Take r
    Take big
    Take Application.ThisWorkbook.Worksheets("S1")
    Take Application.ThisWorkbook
    Take coll
    Take noObj
    HAll = n
End Function

' One traced boundary every object crosses, so each shape pays the same describe
' cost on the same path.
Private Function Take(ByVal v As Variant) As Double
    Take = 1
End Function
'@
     }
     Cells=@{
        'A1' = '1'
        'B1' = '2'
        'C1' = '3'
     }
     FillFormula=@{ Range='E1:E60'; Formula='=HAll(ROW(),$A$1:$C$1,$A:$A)' }
     Trigger=@{ Kind='Calc' }
     Expect={ param($t)
        # The tracer must not have faulted: an opening breaker means repeated faults, which is how a
        # bad COM call would present.
        if ([int]$t.hookFaults -gt 0) {
            return "$($t.hookFaults) hook fault(s) while describing objects" }

        $entries = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.function -eq 'Take' })
        if ($entries.Count -lt 240) {
            return "only $($entries.Count) describe(s) -- the hammer did not land" }

        $vals = @($entries | ForEach-Object {
            if ($_.args -match '^a1:[^=]*=(.+)$') { $Matches[1] } else { '' } })

        # Every one described, not just the first: a describer that degrades (a leaked handle) shows
        # up as later rows falling back to the bare address.
        $bare = @($vals | Where-Object { $_ -match '^object@0x' })
        if ($bare.Count -gt 0) {
            return "$($bare.Count) of $($vals.Count) fell back to a bare address -- the describer degraded under repetition" }

        # The ceiling held every time: a whole column's contents are never read.
        $over = @($vals | Where-Object { $_ -match '^Range@0x[0-9A-F]+\(.*![A-Z]+:[A-Z]+\)=' })
        if ($over.Count -gt 0) {
            return "$($over.Count) whole-column range(s) had their contents read" }

        # And every shape ran, so a silent change of path cannot pass as clean.
        $shapes = @{
            'range'     = @($vals | Where-Object { $_ -match '^Range@0x[0-9A-F]+\(.*!\$?[A-Z]+\$?\d+:' }).Count
            'wholecol'  = @($vals | Where-Object { $_ -match '^Range@0x[0-9A-F]+\(.*![A-Z]+:[A-Z]+\)$' }).Count
            'worksheet' = @($vals | Where-Object { $_ -match '^Worksheet@0x' }).Count
            'workbook'  = @($vals | Where-Object { $_ -match '^Workbook@0x' }).Count
            'collection'= @($vals | Where-Object { $_ -match '^Collection@0x' }).Count
            'nothing'   = @($vals | Where-Object { $_ -eq 'Nothing' }).Count
        }
        foreach ($k in $shapes.Keys) {
            if ($shapes[$k] -lt 40) {
                return ("$k ran only $($shapes[$k]) time(s): " +
                        (($shapes.Keys | ForEach-Object { "$_=$($shapes[$_])" }) -join ' ')) }
        }
        $null }
     Why='every path through the COM describer, several hundred times in one pass:
          what is on trial is not the values but whether calling Excel from a hook
          leaks, faults or degrades under repetition' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
