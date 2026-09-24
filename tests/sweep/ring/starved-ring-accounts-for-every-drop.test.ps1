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
    Check 'the-csv-holds-only-real-events' ($r.EntryExit.Count -eq $r.Rows.Count) `
          ("$($r.Rows.Count - $r.EntryExit.Count) row(s) are neither entry nor exit: " +
           (@($r.Rows | Where-Object { $_.kind -ne 'entry' -and $_.kind -ne 'exit' } | ForEach-Object { $_.kind }) -join ','))
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
