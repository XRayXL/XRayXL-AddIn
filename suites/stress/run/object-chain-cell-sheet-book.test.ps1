$case = @{ Name='object-chain-cell-sheet-book'
     Modules=@{
       'M'=@'
' AN OBJECT WALKED UP ITS OWN PARENTAGE, through three traced calls:
'
'   Chain      selects B2 and takes Application.ActiveCell   -> a Range
'   SheetOf    given that Range, returns r.Parent            -> a Worksheet
'   BookOf     given that Worksheet, returns ws.Parent       -> a Workbook
'
' Each hop crosses the boundary TWICE -- as an argument going in and as a return
' coming out -- so the same object must read identically in both columns, and the
' ADDRESS must let a reader join one call's result to the next call's argument.
' That join is what an address is for, and it is what this case really guards.
'
' NO SCAFFOLDING IN THE TRACE. `Chain` is a macro, so it can WRITE its oracle to
' cells; the driver reads them after disarm. An earlier version called a traced
' helper to get VBA's account into the trace and spent eight rows of fourteen
' doing it.
Public Function SheetOf(ByVal r As Variant) As Variant
    Set SheetOf = r.Parent
End Function

Public Function BookOf(ByVal ws As Variant) As Variant
    Set BookOf = ws.Parent
End Function

Public Sub Chain()
    Dim c As Variant, ws As Variant, wb As Variant, sh As Object
    Set sh = Application.ThisWorkbook.Worksheets("S1")
    sh.Select
    sh.Range("B2").Select
    Set c = Application.ActiveCell
    Set ws = SheetOf(c)
    Set wb = BookOf(ws)

    ' VBA's own account, written where the trace is not.
    sh.Range("E1").Value = TypeName(c)
    sh.Range("E2").Value = TypeName(ws)
    sh.Range("E3").Value = TypeName(wb)
    sh.Range("E4").Value = c.Address(False, False, 1, True)
    sh.Range("E5").Value = ws.Name
    sh.Range("E6").Value = wb.Name
End Sub
'@
     }
     Cells=@{ 'B2' = '7' }
     Trigger=@{ Kind='Run'; Name='Chain' }
     Expect={ param($t)
        $addrOf = { param($s) if ($s -match '@0x([0-9A-Fa-f]+)') { $Matches[1] } else { '' } }
        $entries = @($t.rows | Where-Object { $_.kind -eq 'entry' -and $_.source -eq 'VBA' })
        $exits   = @($t.rows | Where-Object { $_.kind -eq 'exit'  -and $_.source -eq 'VBA' })

        foreach ($k in @('E1','E2','E3','E4','E5','E6')) {
            if (-not $t.cells.ContainsKey($k)) { return "the macro did not record its oracle in $k" }
        }
        if ([string]$t.cells['E1'] -ne 'Range')     { return "VBA called the ActiveCell a '$($t.cells['E1'])'" }
        if ([string]$t.cells['E2'] -ne 'Worksheet') { return "VBA called the parent a '$($t.cells['E2'])'" }
        if ([string]$t.cells['E3'] -ne 'Workbook')  { return "VBA called the grandparent a '$($t.cells['E3'])'" }

        # ---- hop 1: a Range in, a Worksheet out -----------------------------
        $se = @($entries | Where-Object { $_.function -eq 'SheetOf' })
        $sx = @($exits   | Where-Object { $_.function -eq 'SheetOf' })
        if ($se.Count -eq 0 -or $sx.Count -eq 0) { return 'SheetOf was not traced on both sides' }
        $rangeIn  = ArgOf $se[0] 1
        $sheetOut = [string]$sx[0].ret
        if ($rangeIn -notmatch "^Range@0x[0-9A-F]+\('?\[") {
            return "SheetOf's argument is not a described Range: [$rangeIn]" }
        # THE ADDRESS IS EXCEL'S OWN, character for character -- the macro asked
        # Range.Address with the same arguments the tracer uses, quoting and all.
        $vbaAddr = [string]$t.cells['E4']
        if ($rangeIn -notmatch ([regex]::Escape($vbaAddr))) {
            return "the Range reads [$rangeIn]; VBA says its address is '$vbaAddr'" }
        if ($sheetOut -notmatch '^Worksheet@0x[0-9A-F]+\(') {
            return "SheetOf returned [$sheetOut], expected a described Worksheet" }
        if ($sheetOut -notmatch ([regex]::Escape([string]$t.cells['E5']))) {
            return "the Worksheet reads [$sheetOut]; VBA says its name is '$($t.cells['E5'])'" }

        # ---- hop 2: that same Worksheet in, a Workbook out -------------------
        $be = @($entries | Where-Object { $_.function -eq 'BookOf' })
        $bx = @($exits   | Where-Object { $_.function -eq 'BookOf' })
        if ($be.Count -eq 0 -or $bx.Count -eq 0) { return 'BookOf was not traced on both sides' }
        $sheetIn = ArgOf $be[0] 1
        $bookOut = [string]$bx[0].ret
        if ($sheetIn -notmatch '^Worksheet@0x[0-9A-F]+\(') {
            return "BookOf's argument is not a described Worksheet: [$sheetIn]" }
        if ($bookOut -notmatch '^Workbook@0x[0-9A-F]+\(') {
            return "BookOf returned [$bookOut], expected a described Workbook" }
        if ($bookOut -notmatch ([regex]::Escape([string]$t.cells['E6']))) {
            return "the Workbook reads [$bookOut]; VBA says its name is '$($t.cells['E6'])'" }

        # ---- THE JOIN: one object, two columns, one address ------------------
        # What SheetOf RETURNED is what BookOf was GIVEN. Nothing but the address
        # can establish that, which is why it is never dropped from an object.
        $outAddr = & $addrOf $sheetOut
        $inAddr  = & $addrOf $sheetIn
        if (-not $outAddr -or $outAddr -ne $inAddr) {
            return "the Worksheet returned (@0x$outAddr) is not the one passed on (@0x$inAddr) -- an object cannot be followed between rows" }
        $null }
     Why='ActiveCell -> its parent Worksheet -> its parent Workbook, each hop
          crossing as an argument and as a return: the classes and names are
          Excel''s own, and the ADDRESS joins one call''s result to the next
          call''s argument' }
if ($StretchCollectOnly) { return }
. (Join-Path $PSScriptRoot '..\_driver.ps1')
Invoke-StressCase $case
