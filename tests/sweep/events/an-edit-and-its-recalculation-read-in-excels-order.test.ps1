# An edit's events and the calls it causes read in the order Excel raised them.
# VBA's own WithEvents handlers are the witness: the trace must tell the same sequence.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_witness.ps1')

try {
    $sx = Connect-TestExcel
    $app = $sx.App
    Set-XRaySessionDefaults $sx
    $book = New-WitnessBook $sx 'EventOrder'
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $s = Invoke-XRayArmedSession $sx -Body {
        [void]$app.Run("'$($book.Leaf)'!M.ClearLog")
        $book.Sheet.Range('A1').Value2 = 5
    }
    $witness = @(Get-WitnessLog $sx $book)
    $rows = Get-RowsAfterClear @(Read-TraceRows $sx.ProcId)

    $watched = @('SheetChange', 'SheetCalculate', 'AfterCalculate')
    $traced = @($rows | Where-Object {
        ($_.kind -eq 'event' -and $_.source -eq 'Excel' -and $watched -contains $_.function) -or
        ($_.kind -eq 'entry' -and $_.source -eq 'VBA' -and $_.function -eq 'Twice') } | ForEach-Object { $_.function })
    $seen = @($witness | ForEach-Object { $_.Name })

    Check 'the-edit-raised-every-watched-event' (@($watched | Where-Object { $traced -notcontains $_ }).Count -eq 0 -and $traced -contains 'Twice') `
          ("trace: " + ($traced -join ' > '))
    # The witness may go on past disarm (Excel's background pass), so the trace is its prefix.
    $prefix = ($seen.Count -ge $traced.Count) -and (($seen[0..($traced.Count - 1)] -join '>') -eq ($traced -join '>'))
    Check 'the-trace-tells-excels-order' $prefix ("trace: " + ($traced -join ' > ') + "  witness: " + ($seen -join ' > '))

    $change = @($rows | Where-Object { $_.kind -eq 'event' -and $_.function -eq 'SheetChange' }) | Select-Object -First 1
    $wChange = @($witness | Where-Object { $_.Name -eq 'SheetChange' }) | Select-Object -First 1
    Check 'the-change-names-the-edited-cell' ($change -and $wChange -and $change.callerref -eq $wChange.Where) `
          ("trace callerref '$($change.callerref)', witness '$($wChange.Where)'")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("order: " + ($traced -join ' > '))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
