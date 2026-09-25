# Only the ticked events are recorded. The witness shows the others did fire, so their absence
# from the trace is the setting at work, not a quiet session.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_witness.ps1')

function Invoke-Case($Sx, $Book, [string[]]$On, [string]$Cell) {
    $bad = @(Set-XRayEvents $Sx $On)
    if ($bad.Count) { throw ("setter: " + ($bad -join ' | ')) }
    $s = Invoke-XRayArmedSession $Sx -Body {
        [void]$Sx.App.Run("'$($Book.Leaf)'!M.ClearLog")
        $Book.Sheet.Range('A1').Value2 = [double](Get-Random -Minimum 10 -Maximum 99)
        [void]$Book.Sheet.Range($Cell).Select()      # a new cell each case, or nothing changes
    }
    $rows = Get-RowsAfterClear @(Read-TraceRows $Sx.ProcId)
    [pscustomobject]@{
        Events  = @($rows | Where-Object { $_.kind -eq 'event' -and $_.source -eq 'Excel' } | ForEach-Object { $_.function } | Sort-Object -Unique)
        Witness = @(Get-WitnessLog $Sx $Book | ForEach-Object { $_.Name } | Where-Object { $_ -ne 'Twice' } | Sort-Object -Unique)
    }
}

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    $book = New-WitnessBook $sx 'EventChoice'
    [void](Set-XRayTraceParam $sx 'XLL' 'DEPTH' 'OFF')

    $none = Invoke-Case $sx $book @() 'E5'
    Check 'none-records-no-excel-event' ($none.Events.Count -eq 0 -and $none.Witness.Count -ge 3) `
          ("recorded: $($none.Events -join ','); witness saw: $($none.Witness -join ',')")

    $one = Invoke-Case $sx $book @('SheetSelectionChange') 'F6'
    Check 'one-ticked-records-only-that-one' (($one.Events -join ',') -eq 'SheetSelectionChange' -and $one.Witness -contains 'SheetChange') `
          ("recorded: $($one.Events -join ','); witness saw: $($one.Witness -join ',')")

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails case(s) failed" }
    Complete-Test -Pass -Detail ("witness saw " + ($one.Witness -join ','))
}
catch {
    Complete-Test -Fail -Detail ("exception: " + ($_.Exception.Message -replace '\s+', ' ') +
                              $(if ($_.ScriptStackTrace) { "  at " + ($_.ScriptStackTrace -replace '\s+', ' ') } else { '' }))
}
