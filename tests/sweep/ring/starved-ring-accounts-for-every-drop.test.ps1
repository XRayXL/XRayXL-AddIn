# A starved ring accounts for every drop, out of band: written + dropped == opened + closed.
# Fat rows starve the 16 KB minimum ring deterministically: one ~9 KB row nearly fills it, and the
# next arrives before the drain's next pass.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_driver.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    New-XRayMacroBook $sx 'RingStarve' @(@{ Kind = 1; Name = 'RingStarveCase'; Code = $RingFatModule })

    # 16 KB holds exactly one fat row, so under DROP the ring laps and drops
    $r = Invoke-RingCase $sx -Leaf (Get-XRayMacroBook).Leaf -RenderArgs $true -BufferSize '16K' -WhenFull 'DROP'
    Write-Output ("rows=$($r.Rows.Count) entry/exit=$($r.EntryExit.Count) disarmDrops=$($r.DisarmDrops) lineDrops=$($r.LineDrops) inputHoles=$($r.Holes) framesOpened=$($r.Opened) framesClosed=$($r.Closed) expected=$($r.Expected)")

    Check 'buffer-set-echoes-ring' ($r.BufferEcho -match 'ring') "echo: $($r.BufferEcho)"
    Check 'the-ring-actually-dropped' ($r.DisarmDrops -gt 0) `
          "Disarm returned $($r.DisarmDrops) -- the ring was not starved (raise the driver count)"
    # loss is reported out of band, so nothing in the CSV may stand in for it
    $marker = @($r.Rows | Where-Object { $_.kind -ne 'entry' -and $_.kind -ne 'exit' -and $_.kind -ne 'event' })
    Check 'the-csv-holds-only-real-events' ($marker.Count -eq 0) `
          ("$($marker.Count) row(s) are neither a call nor an event: " + (@($marker | ForEach-Object { $_.kind }) -join ','))
    # The session's own rows survive a starved ring, or the trace would read as a crash.
    $all = @(Read-TraceRows $sx.ProcId)
    Check 'the-disarm-row-survives-drop' ($all.Count -and $all[-1].function -eq 'disarm') "last row: $($all[-1].kind) $($all[-1].function)"
    # dropping everything would still reconcile, so require some rows
    Check 'the-starved-ring-still-delivered-rows' ($r.EntryExit.Count -gt 0) 'no rows were written at all'
    Check 'drops-are-locatable-via-input-holes' ($r.Holes -gt 0) `
          "input-column holes=$($r.Holes) -- a dropped row must leave a hole in the input sequence"
    Check 'disarm-return-agrees-with-the-log-line' ($r.DisarmDrops -eq $r.LineDrops) `
          "Disarm returned $($r.DisarmDrops), disarm line said $($r.LineDrops)"
    Check 'nothing-vanished-unaccounted-for' (($r.Opened -ge 0) -and ($r.EntryExit.Count + $r.DisarmDrops -eq $r.Expected)) `
          "written $($r.EntryExit.Count) + dropped $($r.DisarmDrops) = $($r.EntryExit.Count + $r.DisarmDrops), expected $($r.Expected)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails check(s) failed" }
    Complete-Test -Pass -Detail ("starved: wrote $($r.EntryExit.Count), dropped $($r.DisarmDrops), reconciled to $($r.Expected)")
}
catch {
    Complete-Test -Fail -Detail ("threw: " + $_.Exception.Message)
}
