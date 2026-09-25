# An event's row names what VBA's own handler was given: its sheet, and the range it concerns.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_witness.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $book = New-WitnessBook $sx 'EventParams'
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')
    $bad = @(Set-XRayEvents $sx ($script:CalcEvents + 'SheetSelectionChange'))
    if ($bad.Count) { Complete-Test -Fail -Detail ("setter: " + ($bad -join ' | ')) }

    $s = Invoke-XRayArmedSession $sx -Body {
        [void]$app.Run("'$($book.Leaf)'!M.ClearLog")
        $book.Sheet.Range('A1').Value2 = 7
        [void]$book.Sheet.Range('C3:D4').Select()
    }
    $witness = @(Get-WitnessLog $sx $book)
    $rows = Get-RowsAfterClear @(Read-TraceRows $sx.ProcId)

    foreach ($name in 'SheetChange', 'SheetSelectionChange') {
        $w = @($witness | Where-Object { $_.Name -eq $name }) | Select-Object -First 1
        $r = @($rows | Where-Object { $_.kind -eq 'event' -and $_.function -eq $name }) | Select-Object -First 1
        if (-not $w -or -not $r) {
            Check "$name-was-recorded" $false ("witness '$($w.Where)', row '$($r.args)'")
            continue
        }
        $sheetWhere = ($w.Where -replace '!.*$', '')
        Check "$name-callerref-is-the-target" ($r.callerref -eq $w.Where) "row '$($r.callerref)', witness '$($w.Where)'"
        # Contains, not -like: a workbook name's brackets are a wildcard set to -like.
        $args_ = [string]$r.args
        Check "$name-sh-is-the-sheet" ($args_.StartsWith('Sh:Object=Worksheet@0x') -and $args_.Contains("($sheetWhere) ")) "args '$args_'"
        Check "$name-target-is-the-range" ($args_.Contains(' Target:Range=Range@0x') -and $args_.Contains("($($w.Where))")) "args '$args_'"
    }

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("SheetSelectionChange: " + (@($rows | Where-Object { $_.function -eq 'SheetSelectionChange' })[0].args))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
