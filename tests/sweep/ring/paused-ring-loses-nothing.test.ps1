# A paused ring loses nothing, even when starved: BUFFERWHENFULL=PAUSE makes the traced thread wait
# for the drain instead of dropping. Pauses must be above zero, or no backpressure was tested.
. (Join-Path $PSScriptRoot '..\..\..\StretchXL\TestKit.ps1')
. (Join-Path $PSScriptRoot '..\_xray_common.ps1')
. (Join-Path $PSScriptRoot '_driver.ps1')

try {
    $sx = Connect-TestExcel
    Set-XRaySessionDefaults $sx
    New-XRayMacroBook $sx 'RingPause' @(@{ Kind = 1; Name = 'RingPauseCase'; Code = $RingFatModule })

    # the starving ring of the DROP case, under the opposite policy
    $r = Invoke-RingCase $sx -Leaf (Get-XRayMacroBook).Leaf -RenderArgs $true -BufferSize '16K' -WhenFull 'PAUSE'
    Write-Output ("entry/exit=$($r.EntryExit.Count) disarmDrops=$($r.DisarmDrops) lineDrops=$($r.LineDrops) inputHoles=$($r.Holes) pauses=$($r.Pauses) framesOpened=$($r.Opened) framesClosed=$($r.Closed) expected=$($r.Expected)")

    Check 'buffer-set-echoes-ring' ($r.BufferEcho -match 'ring') "echo: $($r.BufferEcho)"
    Check 'whenfull-set-echoes-pause' ($r.WhenFullEcho -match 'PAUSE') "echo: $($r.WhenFullEcho)"
    Check 'the-buffer-actually-paused' ($r.Pauses -gt 0) `
          "pauses=$($r.Pauses) -- the ring never filled (raise the driver count)"
    Check 'nothing-dropped-under-pause' (($r.DisarmDrops -eq 0) -and ($r.LineDrops -eq 0)) `
          "Disarm returned $($r.DisarmDrops), the disarm line said $($r.LineDrops)"
    Check 'no-holes-in-the-producer-sequence' ($r.Holes -eq 0) `
          "$($r.Holes) hole(s) in the input column -- PAUSE must not lose a row"
    Check 'every-frame-is-in-the-file' (($r.Opened -ge 0) -and ($r.EntryExit.Count -eq $r.Expected)) `
          "file entry/exit=$($r.EntryExit.Count) vs framesOpened+framesClosed=$($r.Expected)"

    $checkFails = Get-XRayCheckFailures
    if ($checkFails) { Complete-Test -Fail -Detail "$checkFails check(s) failed" }
    Complete-Test -Pass -Detail ("paused: wrote $($r.EntryExit.Count), dropped 0, paused $($r.Pauses) time(s)")
}
catch {
    Complete-Test -Fail -Detail ("threw: " + $_.Exception.Message)
}
